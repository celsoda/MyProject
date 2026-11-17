/**
 * Local smoke test without Azure credentials (MOCK_GRAPH=1).
 * Exercises HTTP /v1/mail/messages and the shared Connector used by MCP tools.
 */
import { loadConfig } from "../src/config.js";
import { Connector } from "../src/connector.js";
import { createHttpApp } from "../src/httpApp.js";

async function main(): Promise<void> {
  process.env.MOCK_GRAPH = "1";
  delete process.env.AZURE_CLIENT_ID;

  const config = loadConfig();
  if (!config.mockGraph) {
    throw new Error("Expected MOCK_GRAPH=1 for smoke");
  }

  const connector = new Connector(config);
  const login = await connector.startLogin();
  if (!login.userCode) throw new Error("expected mock device code");

  await connector.waitLogin();
  const status = connector.authStatus();
  if (!status.linked) throw new Error("expected linked after mock login");

  const messages = await connector.listMail(5);
  if (messages.length < 1) throw new Error("expected mock messages");

  const sent = await connector.sendMail({
    to: ["demo@example.com"],
    subject: "smoke",
    bodyText: "hello",
  });
  if (!sent.ok) throw new Error("expected send ok");

  const app = createHttpApp(connector);
  const res = await app.request("http://local/v1/mail/messages?limit=2");
  if (!res.ok) {
    throw new Error(`HTTP list failed: ${res.status} ${await res.text()}`);
  }
  const json = (await res.json()) as { messages: unknown[] };
  if (!Array.isArray(json.messages) || json.messages.length < 1) {
    throw new Error("HTTP response missing messages");
  }

  console.log(
    JSON.stringify(
      {
        ok: true,
        mode: "mock",
        auth: status,
        messageCount: messages.length,
        httpCount: json.messages.length,
        sampleSubject: messages[0]?.subject,
      },
      null,
      2,
    ),
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
