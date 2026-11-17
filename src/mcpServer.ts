/**
 * Claude Desktop MCP server (stdio).
 * LEFT: Claude Desktop tools → RIGHT: Microsoft Graph via Connector.
 * Microsoft tokens never leave this process / encrypted local cache.
 */
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { loadConfig } from "./config.js";
import { Connector } from "./connector.js";

const config = loadConfig();
const connector = new Connector(config);

const server = new McpServer({
  name: "outlook-connector",
  version: "0.1.0",
});

function text(data: unknown) {
  return {
    content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }],
  };
}

server.tool(
  "auth_status",
  "Show whether Outlook/Microsoft is linked for this connector (no tokens returned).",
  {},
  async () => text(connector.authStatus()),
);

server.tool(
  "outlook_login",
  "Start Microsoft device-code login (public client, no client secret). Returns a user code and verification URL. Open the URL in a browser, enter the code, consent Mail.Read. Then call auth_status or list_mail. For MOCK_GRAPH=1, login is simulated immediately.",
  {},
  async () => {
    const info = await connector.startLogin();
    // Wait in background so Claude can show the code; completion updates cache.
    void connector.waitLogin().catch((err: unknown) => {
      console.error(
        "Device-code wait failed:",
        err instanceof Error ? err.message : err,
      );
    });
    return text({
      instructions:
        "Open verificationUri in a browser, enter userCode, sign in to Microsoft, and approve Mail.Read + User.Read. Then call list_mail.",
      ...info,
    });
  },
);

server.tool(
  "outlook_logout",
  "Clear the locally stored Microsoft token cache for this connector.",
  {},
  async () => {
    await connector.logout();
    return text({ ok: true });
  },
);

server.tool(
  "list_mail",
  "List recent Inbox messages for the linked Microsoft account (delegated Graph Mail.Read). Claude never receives Microsoft access tokens.",
  {
    limit: z
      .number()
      .int()
      .min(1)
      .max(50)
      .optional()
      .describe("Max messages to return (default 10)"),
  },
  async ({ limit }) => {
    const messages = await connector.listMail(limit ?? 10);
    return text({ messages });
  },
);

server.tool(
  "send_mail",
  "Send email as the linked Microsoft user via delegated Graph Mail.Send. In MOCK_GRAPH mode, send is simulated and no message leaves the machine.",
  {
    to: z.array(z.string().email()).min(1).describe("Recipient addresses"),
    subject: z.string().min(1),
    bodyText: z.string().describe("Plain-text body"),
  },
  async ({ to, subject, bodyText }) => {
    const result = await connector.sendMail({ to, subject, bodyText });
    return text(result);
  },
);

async function main(): Promise<void> {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(
    `outlook-connector MCP ready (mode=${config.mockGraph ? "mock" : "live"})`,
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
