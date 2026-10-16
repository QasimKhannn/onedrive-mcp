const { graphGet, graphGetRaw } = require("../graph");
const {
  createTextResult,
  formatBytes,
  formatDate,
  userError
} = require("../response");
const mammoth = require("mammoth");
const { PDFParse } = require("pdf-parse");
const XLSX = require("xlsx");
const JSZip = require("jszip");
const { XMLParser } = require("fast-xml-parser");

const MAX_BYTES = 20 * 1024 * 1024;
const DEFAULT_MAX_CHARS = 12000;
const SUPPORTED_TYPES = ["docx", "pdf", "xlsx", "xls", "pptx", "txt", "md", "csv", "json"];

function trimText(text, maxChars) {
  if (!text) {
    return "";
  }

  if (text.length <= maxChars) {
    return text;
  }

  return `${text.slice(0, maxChars)}\n\n...(truncated)...`;
}

function readMaxChars(value) {
  if (value === undefined || value === null) {
    return DEFAULT_MAX_CHARS;
  }

  if (typeof value !== "number" || Number.isNaN(value) || !Number.isInteger(value) || value < 1) {
    throw userError("Please provide maxChars as a whole number greater than 0.", {
      code: "invalid_parameter",
      statusCode: 400
    });
  }

  return value;
}

async function parseDocx(buffer) {
  try {
    const result = await mammoth.extractRawText({ buffer });
    return result.value;
  } catch (error) {
    throw userError("I couldn't read the text from this Word document.", {
      code: "file_parse_failed",
      statusCode: 422,
      internalMessage: error.message
    });
  }
}

async function parsePdf(buffer) {
  let parser;

  try {
    parser = new PDFParse({ data: buffer });
    const result = await parser.getText();
    return result.text;
  } catch (error) {
    throw userError("I couldn't read the text from this PDF.", {
      code: "file_parse_failed",
      statusCode: 422,
      internalMessage: error.message
    });
  } finally {
    if (parser) {
      await parser.destroy();
    }
  }
}

function parseXlsx(buffer) {
  try {
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const sections = [];

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      const sheetText = rows
        .map((row) => row.map((cell) => (cell === undefined || cell === null ? "" : String(cell))).join("\t"))
        .join("\n");

      sections.push(`### Sheet: ${sheetName}\n${sheetText}`);
    }

    return sections.join("\n\n");
  } catch (error) {
    throw userError("I couldn't read the data from this Excel file.", {
      code: "file_parse_failed",
      statusCode: 422,
      internalMessage: error.message
    });
  }
}

async function parsePptx(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const parser = new XMLParser({
      ignoreAttributes: true,
      removeNSPrefix: true
    });
    const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path));
    const sections = [];

    for (const slidePath of slides) {
      const xml = await zip.file(slidePath).async("string");
      const parsed = parser.parse(xml);
      const textRuns = [];

      const walk = (node) => {
        if (!node || typeof node !== "object") {
          return;
        }

        for (const [key, value] of Object.entries(node)) {
          if (key === "t" && typeof value === "string") {
            textRuns.push(value);
          } else if (Array.isArray(value)) {
            value.forEach(walk);
          } else if (typeof value === "object") {
            walk(value);
          }
        }
      };

      walk(parsed);

      const slideNumber = slidePath.match(/slide(\d+)/i)?.[1] || "?";
      sections.push(`### Slide ${slideNumber}\n${textRuns.join(" ")}`);
    }

    return sections.join("\n\n");
  } catch (error) {
    throw userError("I couldn't read the text from this PowerPoint file.", {
      code: "file_parse_failed",
      statusCode: 422,
      internalMessage: error.message
    });
  }
}

function parsePlain(buffer) {
  return buffer.toString("utf8");
}

async function extractText(buffer, extension) {
  if (extension === "docx") {
    return parseDocx(buffer);
  }

  if (extension === "pdf") {
    return parsePdf(buffer);
  }

  if (extension === "xlsx" || extension === "xls") {
    return parseXlsx(buffer);
  }

  if (extension === "pptx") {
    return parsePptx(buffer);
  }

  return parsePlain(buffer);
}

const tools = [
  {
    name: "get_onedrive_item",
    description: "Show a user-friendly summary of a specific OneDrive file or folder by item ID, including type, size, location, and update details.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "The OneDrive item ID of the file or folder."
        }
      },
      required: ["itemId"],
      additionalProperties: false
    }
  },
  {
    name: "read_file_content",
    description: "Download a OneDrive file and extract readable text from supported formats such as Word, PDF, Excel, PowerPoint, text, CSV, Markdown, and JSON.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "The OneDrive item ID of the file to read."
        },
        maxChars: {
          type: "number",
          description: "Optional maximum number of characters to return from the extracted text."
        }
      },
      required: ["itemId"],
      additionalProperties: false
    }
  }
];

const handlers = {
  get_onedrive_item: async ({ accessToken, params }) => {
    const { itemId } = params;

    const item = await graphGet(
      accessToken,
      `/me/drive/items/${itemId}`
    );

    return createTextResult([
      "Here is the OneDrive item you requested.",
      "",
      `Name: ${item.name || "Untitled"}`,
      `Type: ${item.folder ? "Folder" : item.file ? "File" : "Item"}`,
      `Item ID: ${item.id || "Not available"}`,
      `Size: ${item.folder ? `${item.folder.childCount || 0} items` : formatBytes(item.size)}`,
      `Last modified: ${formatDate(item.lastModifiedDateTime)}`,
      `Last modified by: ${item.lastModifiedBy?.user?.displayName || "Not available"}`,
      `Path: ${item.parentReference?.path || "Not available"}`,
      item.webUrl ? `Open in browser: ${item.webUrl}` : null,
      item.file?.mimeType ? `File type: ${item.file.mimeType}` : null
    ].filter(Boolean).join("\n"));
  },

  read_file_content: async ({ accessToken, params }) => {
    const itemId = typeof params?.itemId === "string" ? params.itemId.trim() : "";
    const maxChars = readMaxChars(params?.maxChars);

    if (!itemId) {
      throw userError("Please provide the OneDrive item ID of the file you want to read.", {
        code: "missing_parameter",
        statusCode: 400
      });
    }

    const item = await graphGet(accessToken, `/me/drive/items/${itemId}`);

    if (item.folder) {
      throw userError("That OneDrive item is a folder, so there is no file content to read.", {
        code: "not_a_file",
        statusCode: 400,
        hint: "Choose a file item instead of a folder."
      });
    }

    const name = item.name || "file";
    const extension = (name.split(".").pop() || "").toLowerCase();

    if (!SUPPORTED_TYPES.includes(extension)) {
      throw userError(
        `This file type is not supported for text extraction yet. Supported formats are ${SUPPORTED_TYPES.map((type) => `.${type}`).join(", ")}.`,
        {
          code: "unsupported_file_type",
          statusCode: 400
        }
      );
    }

    if ((item.size || 0) > MAX_BYTES) {
      throw userError(
        `This file is too large to read here. The limit is ${(MAX_BYTES / 1024 / 1024).toFixed(0)} MB.`,
        {
          code: "file_too_large",
          statusCode: 400
        }
      );
    }

    const buffer = await graphGetRaw(accessToken, `/me/drive/items/${itemId}/content`);
    const extractedText = await extractText(buffer, extension);
    const trimmedText = trimText(extractedText, maxChars);

    return createTextResult(
      [
        `File: ${name}`,
        `Type: .${extension}`,
        `Size: ${formatBytes(item.size)}`,
        item.webUrl ? `Open in browser: ${item.webUrl}` : null,
        "",
        trimmedText || "I could open the file, but there was no readable text to extract."
      ].filter((line) => line !== null).join("\n"),
      {
        structuredContent: {
          file: {
            id: item.id,
            name,
            extension,
            size: item.size,
            webUrl: item.webUrl
          },
          extractedText: trimmedText,
          truncated: extractedText.length > maxChars
        }
      }
    );
  }
};

module.exports = { tools, handlers };
