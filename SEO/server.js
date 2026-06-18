import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

const server = new Server(
  {
    name: "wordstat-mcp",
    version: "1.0.0"
  },
  {
    capabilities: {
      tools: {}
    }
  }
);

server.setRequestHandler("tools/list", async () => ({
  tools: [
    {
      name: "hello_wordstat",
      description: "Тестовый инструмент",
      inputSchema: {
        type: "object",
        properties: {}
      }
    }
  ]
}));

server.setRequestHandler("tools/call", async (request) => {
  if (request.params.name === "hello_wordstat") {
    return {
      content: [
        {
          type: "text",
          text: "MCP работает!"
        }
      ]
    };
  }
});

const transport = new StdioServerTransport();
await server.connect(transport);
