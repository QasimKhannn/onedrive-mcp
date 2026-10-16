const axios = require("axios");

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

async function graphGet(accessToken, path) {
  const res = await axios.get(`${GRAPH_BASE}${path}`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  return res.data;
}

async function graphPost(accessToken, path, body) {
  const res = await axios.post(`${GRAPH_BASE}${path}`, body, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  return res.data;
}

async function graphPatch(accessToken, path, body) {
  const res = await axios.patch(`${GRAPH_BASE}${path}`, body, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  return res.data;
}

async function graphDelete(accessToken, path) {
  await axios.delete(`${GRAPH_BASE}${path}`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  return { success: true };
}

async function graphGetRaw(accessToken, path) {
  const res = await axios.get(`${GRAPH_BASE}${path}`, {
    headers: { Authorization: `Bearer ${accessToken}` },
    responseType: "arraybuffer",
  });

  return Buffer.from(res.data);
}

module.exports = {
  graphGet,
  graphPost,
  graphPatch,
  graphDelete,
  graphGetRaw,
};