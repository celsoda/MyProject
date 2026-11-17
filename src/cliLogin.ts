import { loadConfig } from "./config.js";
import { Connector } from "./connector.js";

async function main(): Promise<void> {
  const config = loadConfig();
  const connector = new Connector(config);
  console.error("Starting Microsoft device-code login…");
  const status = await connector.loginInteractive((message) => {
    console.error("\n" + message + "\n");
  });
  console.log(JSON.stringify(status, null, 2));
  console.error("Login complete. You can use Claude Desktop MCP tools or HTTP /v1/mail/messages.");
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : err);
  process.exit(1);
});
