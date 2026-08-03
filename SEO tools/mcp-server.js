import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { getWordstat } from "./wordstat.js";
import { saveToCSV } from "./save-seo.js";

const server = new Server(
  { name: "seo-mcp", version: "1.1.0" },
  { capabilities: { tools: {} } }
);

server.setRequestHandler("tools/list", async () => {
  return {
    tools: [
      {
        name: "ping",
        description: "check mcp",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "wordstat",
        description: "get keyword ideas from yandex wordstat",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string" },
          },
          required: ["query"],
        },
      },
      {
        name: "save_seo",
        description: "save wordstat data to seo folder",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string" },
            data: { type: "array" },
          },
          required: ["query", "data"],
        },
      },
    ],
  };
});

server.setRequestHandler("tools/call", async (req) => {
  if (req.params.name === "ping") {
    return {
      content: [{ type: "text", text: "pong ✅ MCP OK" }],
    };
  }

  if (req.params.name === "wordstat") {
    const result = await getWordstat(req.params.arguments.query);

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  }

  if (req.params.name === "save_seo") {
    const file = saveToCSV(req.params.arguments.query, req.params.arguments.data);

    return {
      content: [
        {
          type: "text",
          text: "saved: " + file,
        },
      ],
    };
  }
});

const transport = new StdioServerTransport();
server.connect(transport);
