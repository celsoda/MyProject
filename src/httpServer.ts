import { serve } from "@hono/node-server";
import { loadConfig } from "./config.js";
import { Connector } from "./connector.js";
import { createHttpApp } from "./httpApp.js";

const config = loadConfig();
const connector = new Connector(config);
const app = createHttpApp(connector, config.CONNECTOR_API_KEY);

serve({ fetch: app.fetch, port: config.PORT }, (info) => {
  console.error(
    `Outlook connector HTTP listening on http://127.0.0.1:${info.port} (mode=${config.mockGraph ? "mock" : "live"})`,
  );
  console.error(
    "Claude Desktop should prefer the MCP stdio entry (`npm run mcp`), not this HTTP server.",
  );
});
