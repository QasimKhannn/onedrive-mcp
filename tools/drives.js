const { graphGet } = require("../graph");
const { createTextResult, formatDriveItemList } = require("../response");

const tools = [
  {
    name: "list_onedrive_root",
    description: "Browse the top level of the signed-in user's OneDrive and return a readable list of files and folders.",
    inputSchema: {
      type: "object",
      properties: {},
      additionalProperties: false
    },
  },
  {
    name: "list_onedrive_children",
    description: "Browse the contents of a specific OneDrive folder by item ID and return a readable list of what is inside.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "The OneDrive item ID of the folder to open."
        }
      },
      required: ["itemId"],
      additionalProperties: false
    }
  }
];

const handlers = {

  list_onedrive_root: async ({ accessToken }) => {

    const data = await graphGet(accessToken, `/me/drive/root/children`);

    return createTextResult(
      formatDriveItemList(
        "Top-level OneDrive items",
        data.value,
        "The OneDrive root is empty."
      )
    );
  },

  list_onedrive_children: async ({ accessToken, params }) => {

    const { itemId } = params;

    const data = await graphGet(
      accessToken,
      `/me/drive/items/${itemId}/children`
    );

    return createTextResult(
      formatDriveItemList(
        `Contents of folder ${itemId}`,
        data.value,
        "This folder does not contain any items."
      )
    );
  }

};

module.exports = { tools, handlers };
