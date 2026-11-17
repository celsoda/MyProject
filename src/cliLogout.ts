import { loadConfig } from "./config.js";
import { Connector } from "./connector.js";

async function main(): Promise<void> {
  const config = loadConfig();
  const connector = new Connector(config);
  await connector.logout();
  console.log(JSON.stringify({ ok: true }, null, 2));
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : err);
  process.exit(1);
});
