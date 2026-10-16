class ToolError extends Error {
  constructor(publicMessage, options = {}) {
    super(options.internalMessage || publicMessage);
    this.name = "ToolError";
    this.publicMessage = publicMessage;
    this.statusCode = options.statusCode || 400;
    this.code = options.code || "tool_error";
    this.hint = options.hint;
  }
}

function createTextResult(text, options = {}) {
  return {
    isError: Boolean(options.isError),
    structuredContent: options.structuredContent || undefined,
    content: [
      {
        type: "text",
        text
      }
    ]
  };
}

function createErrorResult(message, nextSteps = []) {
  const lines = [message];

  if (nextSteps.length) {
    lines.push("");
    lines.push("What you can do next:");
    nextSteps.forEach((step) => lines.push(`- ${step}`));
  }

  return createTextResult(lines.join("\n"), { isError: true });
}

function createStructuredErrorResult(message, options = {}) {
  const parts = [message];

  if (options.hint) {
    parts.push("");
    parts.push(options.hint);
  }

  return createTextResult(parts.join("\n"), {
    isError: true,
    structuredContent: {
      error: {
        code: options.code || "tool_error",
        statusCode: options.statusCode || 400
      }
    }
  });
}

function userError(publicMessage, options = {}) {
  return new ToolError(publicMessage, options);
}

function formatBytes(bytes) {
  if (typeof bytes !== "number" || Number.isNaN(bytes)) {
    return "Not available";
  }

  if (bytes === 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB", "TB"];
  const exponent = Math.min(
    Math.floor(Math.log(bytes) / Math.log(1024)),
    units.length - 1
  );
  const value = bytes / 1024 ** exponent;
  const precision = value >= 100 || exponent === 0 ? 0 : value >= 10 ? 1 : 2;

  return `${value.toFixed(precision)} ${units[exponent]}`;
}

function formatDate(value) {
  if (!value) {
    return "Not available";
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "Not available";
  }

  return date.toISOString();
}

function formatItemKind(item) {
  if (!item || typeof item !== "object") {
    return "Item";
  }

  if (item.folder) {
    return "Folder";
  }

  if (item.file) {
    return "File";
  }

  return "Item";
}

function formatDriveItemLine(item) {
  const kind = formatItemKind(item);
  const size = item?.folder
    ? `${item.folder.childCount || 0} items`
    : formatBytes(item?.size);

  return `- ${kind}: ${item?.name || "Untitled"} | ID: ${item?.id || "Not available"} | Size: ${size} | Modified: ${formatDate(item?.lastModifiedDateTime)}`;
}

function formatDriveItemList(title, items, emptyMessage) {
  const safeItems = Array.isArray(items) ? items : [];

  if (!safeItems.length) {
    return `${title}\n\n${emptyMessage}`;
  }

  return [
    `${title}`,
    "",
    `Found ${safeItems.length} item${safeItems.length === 1 ? "" : "s"}:`,
    ...safeItems.map(formatDriveItemLine)
  ].join("\n");
}

function formatGraphError(error, fallbackMessage) {
  if (error instanceof ToolError) {
    return createStructuredErrorResult(
      error.publicMessage,
      {
        code: error.code,
        statusCode: error.statusCode,
        hint: error.hint
      }
    );
  }

  const status = error?.response?.status;

  if (status === 400) {
    return createErrorResult(
      "I couldn't complete that request because some of the information sent to OneDrive was invalid.",
      ["Check the item ID or search text and try again."]
    );
  }

  if (status === 401) {
    return createErrorResult(
      "Your Microsoft connection has expired, so I couldn't reach OneDrive.",
      ["Sign in again, then retry the request."]
    );
  }

  if (status === 403) {
    return createErrorResult(
      "You do not have permission to access that OneDrive content.",
      ["Try a file or folder you own or have been shared access to."]
    );
  }

  if (status === 404) {
    return createErrorResult(
      "I couldn't find that file or folder in OneDrive.",
      ["Double-check the item ID, or browse/search for it again."]
    );
  }

  if (status === 409) {
    return createErrorResult(
      "OneDrive reported a conflict while processing that request.",
      ["Wait a moment and try again."]
    );
  }

  if (status === 429) {
    return createErrorResult(
      "OneDrive is receiving too many requests right now.",
      ["Wait a few moments, then try again."]
    );
  }

  if (status >= 500) {
    return createErrorResult(
      "Microsoft OneDrive is having trouble right now, so the request could not be completed.",
      ["Try again in a little while."]
    );
  }

  if (error?.code === "ECONNABORTED" || error?.code === "ENOTFOUND") {
    return createErrorResult(
      "I couldn't reach Microsoft OneDrive because of a network issue.",
      ["Check the connection and try again."]
    );
  }

  return createErrorResult(
    fallbackMessage || "Something went wrong while working with OneDrive.",
    ["Try again.", "If the issue keeps happening, reconnect your Microsoft account."]
  );
}

module.exports = {
  ToolError,
  createTextResult,
  createErrorResult,
  createStructuredErrorResult,
  userError,
  formatBytes,
  formatDate,
  formatDriveItemLine,
  formatDriveItemList,
  formatGraphError
};
