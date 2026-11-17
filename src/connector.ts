import { AuthService } from "./auth.js";
import type { AppConfig } from "./config.js";
import { listInboxMessages, MOCK_MESSAGES, sendMail } from "./graph.js";
import type { AuthStatus, DeviceCodeStart, MailMessageSummary } from "./types.js";

/**
 * Shared service used by HTTP API, MCP (Claude Desktop), and CLI.
 * Claude never sees Microsoft tokens — only tool/HTTP results.
 */
export class Connector {
  readonly auth: AuthService;

  constructor(private readonly config: AppConfig) {
    this.auth = new AuthService(config);
  }

  authStatus(): AuthStatus {
    return this.auth.status();
  }

  loginInteractive(onMessage?: (message: string) => void): Promise<AuthStatus> {
    return this.auth.loginInteractive(onMessage);
  }

  startLogin(): Promise<DeviceCodeStart> {
    return this.auth.startDeviceLogin();
  }

  waitLogin(): Promise<AuthStatus> {
    return this.auth.waitForDeviceLogin();
  }

  logout(): Promise<void> {
    return this.auth.logout();
  }

  async listMail(limit = 10): Promise<MailMessageSummary[]> {
    if (this.config.mockGraph) {
      await this.auth.getAccessToken();
      return MOCK_MESSAGES.slice(0, limit);
    }
    const token = await this.auth.getAccessToken();
    return listInboxMessages(token, limit);
  }

  async sendMail(input: {
    to: string[];
    subject: string;
    bodyText: string;
  }): Promise<{ ok: true; mock?: boolean }> {
    if (this.config.mockGraph) {
      await this.auth.getAccessToken();
      return { ok: true, mock: true };
    }
    const token = await this.auth.getAccessToken();
    await sendMail(token, input);
    return { ok: true };
  }
}
