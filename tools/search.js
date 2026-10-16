const { graphGet } = require("../graph");
const { createTextResult, formatDriveItemList } = require("../response");

const tools = [
  {
    name: "search_onedrive",
    description: "Search the signed-in user's OneDrive for files and folders that match a query and return clear, readable results.",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "The file or folder name, keyword, or phrase to search for in OneDrive."
        }
      },
      required: ["query"],
      additionalProperties: false
    }
  }
];

const handlers = {

  search_onedrive: async ({ accessToken, params }) => {

    const { query } = params;

    const data = await graphGet(
      accessToken,
      `/me/drive/root/search(q='${encodeURIComponent(query)}')`
    );

    return createTextResult(
      formatDriveItemList(
        `Search results for "${query}"`,
        data.value,
        "No OneDrive items matched that search."
      )
    );
  }

};

module.exports = { tools, handlers };
