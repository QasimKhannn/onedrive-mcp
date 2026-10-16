require("dotenv").config();
const express = require("express");

const drives = require("./tools/drives");
const files = require("./tools/files");
const search = require("./tools/search");
const users = require("./tools/users");
const { createErrorResult, formatGraphError } = require("./response");

const app = express();
app.use(express.json());

const PORT = process.env.PORT || 8080;

function getBearerToken(req) {
  const auth = req.headers["authorization"] || "";
  return auth.startsWith("Bearer ") ? auth.slice(7) : null;
}

const tools = [
  ...users.tools,
  ...drives.tools,
  ...files.tools,
  ...search.tools
];

const handlers = {
  ...drives.handlers,
  ...files.handlers,
  ...search.handlers,
  ...users.handlers
};

app.post("/mcp", async (req, res) => {

  const { id, method, params } = req.body;
  const accessToken = getBearerToken(req);

  if (method === "initialize") {
    return res.json({
      jsonrpc: "2.0",
      id,
      result: {
        protocolVersion: "2024-11-05",
        capabilities: { tools: {} },
        serverInfo: {
          name: "onedrive-mcp",
          version: "1.0.0"
        }
      }
    });
  }

  if (method === "tools/list") {
    return res.json({
      jsonrpc: "2.0",
      id,
      result: { tools }
    });
  }

  if (!accessToken) {
    if (method === "tools/call") {
      return res.json({
        jsonrpc: "2.0",
        id,
        result: createErrorResult(
          "I couldn't access OneDrive because this session is not signed in.",
          ["Connect your Microsoft account, then try again."]
        )
      });
    }

    return res.status(401).json({
      jsonrpc: "2.0",
      id,
      error: { code: 401, message: "Authentication required" }
    });
  }

  if (method === "tools/call") {

    const toolName = params?.name;
    const args = params?.arguments || {};

    if (!toolName) {
      return res.json({
        jsonrpc: "2.0",
        id,
        result: createErrorResult(
          "I couldn't complete that request because the OneDrive action name was missing.",
          ["Retry the request with a valid OneDrive action."]
        )
      });
    }

    const handler = handlers[toolName];

    if (!handler) {
      return res.json({
        jsonrpc: "2.0",
        id,
        result: createErrorResult(
          "I couldn't complete that request because the OneDrive action was not recognized.",
          ["Try again with a supported OneDrive action."]
        )
      });
    }

    try {
      const result = await handler({
        accessToken,
        params: args
      });

      return res.json({
        jsonrpc: "2.0",
        id,
        result
      });

    } catch (err) {
      return res.json({
        jsonrpc: "2.0",
        id,
        result: formatGraphError(
          err,
          "I ran into a problem while trying to complete that OneDrive request."
        )
      });

    }

  }

  return res.json({
    jsonrpc: "2.0",
    id,
    error: { code: -32601, message: "Method not found" }
  });

});

app.listen(PORT, () =>
  console.log(`OneDrive MCP running on http://localhost:${PORT}`)
);
