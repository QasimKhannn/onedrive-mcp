const { graphGet } = require("../graph");
const { createTextResult, formatBytes } = require("../response");

const tools = [
  {
    name: "get_current_user",
    description: "Show who is currently signed in to Microsoft so the user can confirm which account this OneDrive session is using.",
    inputSchema: {
      type: "object",
      properties: {},
      additionalProperties: false
    }
  },
  {
    name: "get_my_onedrive",
    description: "Show a friendly summary of the signed-in user's OneDrive, including drive details, owner information, and storage usage.",
    inputSchema: {
      type: "object",
      properties: {},
      additionalProperties: false
    }
  },
  {
    name: "get_onedrive_quota",
    description: "Show how much OneDrive storage the signed-in user has used, how much is left, and the current quota status.",
    inputSchema: {
      type: "object",
      properties: {},
      additionalProperties: false
    }
  }
];

const handlers = {

  get_current_user: async ({ accessToken }) => {

    const user = await graphGet(accessToken, "/me");

    return createTextResult([
      "Microsoft account connected successfully.",
      "",
      `Name: ${user.displayName || "Not available"}`,
      `Email: ${user.mail || user.userPrincipalName || "Not available"}`,
      `User ID: ${user.id || "Not available"}`
    ].join("\n"));
  },

  get_my_onedrive: async ({ accessToken }) => {

    const drive = await graphGet(accessToken, "/me/drive");
    const quota = drive.quota || {};

    return createTextResult([
      "Here is a summary of the connected OneDrive.",
      "",
      `Drive name: ${drive.name || "Not available"}`,
      `Drive type: ${drive.driveType || "Not available"}`,
      `Drive ID: ${drive.id || "Not available"}`,
      `Owner: ${drive.owner?.user?.displayName || "Not available"}`,
      `Storage used: ${formatBytes(quota.used)}`,
      `Storage remaining: ${formatBytes(quota.remaining)}`,
      `Storage total: ${formatBytes(quota.total)}`,
      `Quota status: ${quota.state || "Not available"}`,
      drive.webUrl ? `Open in browser: ${drive.webUrl}` : null
    ].filter(Boolean).join("\n"));
  },

  get_onedrive_quota: async ({ accessToken }) => {

    const drive = await graphGet(accessToken, "/me/drive");
    const quota = drive.quota || {};

    return createTextResult([
      "Here is the current OneDrive storage usage.",
      "",
      `Used: ${formatBytes(quota.used)}`,
      `Remaining: ${formatBytes(quota.remaining)}`,
      `Total: ${formatBytes(quota.total)}`,
      `Deleted but recoverable: ${formatBytes(quota.deleted)}`,
      `Quota status: ${quota.state || "Not available"}`
    ].join("\n"));
  }

};

module.exports = { tools, handlers };
