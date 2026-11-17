import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type DeviceCodeRequest,
} from "@azure/msal-node";
import type { AppConfig } from "./config.js";
import {
  clearSession,
  createEncryptedCachePlugin,
  loadSessionMeta,
  saveSessionMeta,
} from "./tokenStore.js";
import type { AuthStatus, DeviceCodeStart } from "./types.js";

type PendingDeviceLogin = {
  startedAt: number;
  info: DeviceCodeStart;
  promise: Promise<AuthenticationResult>;
};

export class AuthService {
  private readonly pca: PublicClientApplication | null;
  private pending: PendingDeviceLogin | null = null;
  private mockLinked = false;

  constructor(private readonly config: AppConfig) {
    if (config.mockGraph) {
      this.pca = null;
      return;
    }

    this.pca = new PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: config.authority,
      },
      cache: {
        cachePlugin: createEncryptedCachePlugin(
          config.DATA_DIR,
          config.TOKEN_ENCRYPTION_KEY,
        ),
      },
    });
  }

  status(): AuthStatus {
    if (this.config.mockGraph) {
      return {
        linked: this.mockLinked,
        mode: "mock",
        username: this.mockLinked ? "mock.user@example.com" : undefined,
        name: this.mockLinked ? "Mock User" : undefined,
      };
    }

    const meta = loadSessionMeta(this.config.DATA_DIR);
    return {
      linked: Boolean(meta?.homeAccountId),
      mode: "live",
      username: meta?.username,
      name: meta?.name,
      homeAccountId: meta?.homeAccountId,
    };
  }

  async logout(): Promise<void> {
    this.pending = null;
    if (this.config.mockGraph) {
      this.mockLinked = false;
      return;
    }
    clearSession(this.config.DATA_DIR);
  }

  /**
   * Interactive CLI login: prints device-code instructions to stderr and waits.
   */
  async loginInteractive(
    onMessage: (message: string) => void = (m) => console.error(m),
  ): Promise<AuthStatus> {
    if (this.config.mockGraph) {
      this.mockLinked = true;
      return this.status();
    }

    const result = await this.acquireByDeviceCode(onMessage);
    this.persistAccount(result.account);
    return this.status();
  }

  /**
   * Start device-code login for HTTP/MCP. Returns codes immediately; completion
   * continues in the background until pollDeviceLogin / getAccessToken.
   */
  async startDeviceLogin(): Promise<DeviceCodeStart> {
    if (this.config.mockGraph) {
      this.mockLinked = true;
      return {
        userCode: "MOCK-CODE",
        verificationUri: "https://example.com/mock-login",
        message:
          "MOCK_GRAPH=1: login simulated. Call list_mail / GET /v1/mail/messages next.",
        expiresIn: 900,
        interval: 5,
      };
    }

    if (this.pending) {
      return this.pending.info;
    }

    let resolveInfo!: (info: DeviceCodeStart) => void;
    const infoPromise = new Promise<DeviceCodeStart>((resolve) => {
      resolveInfo = resolve;
    });

    const request: DeviceCodeRequest = {
      scopes: this.config.scopes,
      deviceCodeCallback: (response) => {
        const info: DeviceCodeStart = {
          userCode: response.userCode,
          verificationUri: response.verificationUri,
          message: response.message,
          expiresIn: response.expiresIn,
          interval: response.interval,
        };
        resolveInfo(info);
      },
    };

    const promise = this.pca!.acquireTokenByDeviceCode(request).then(
      (result) => {
        if (!result) {
          throw new Error("Device code login returned no result");
        }
        this.persistAccount(result.account);
        this.pending = null;
        return result;
      },
      (err: unknown) => {
        this.pending = null;
        throw err;
      },
    );

    // Prime callback by starting acquire; info arrives via deviceCodeCallback.
    // Attach a no-op catch so early rejection does not become unhandled.
    promise.catch(() => undefined);

    const info = await infoPromise;
    this.pending = { startedAt: Date.now(), info, promise };
    return info;
  }

  async waitForDeviceLogin(): Promise<AuthStatus> {
    if (this.config.mockGraph) {
      this.mockLinked = true;
      return this.status();
    }
    if (!this.pending) {
      return this.status();
    }
    await this.pending.promise;
    return this.status();
  }

  async getAccessToken(): Promise<string> {
    if (this.config.mockGraph) {
      if (!this.mockLinked) {
        throw new Error("Not logged in. Call auth login / outlook_login first.");
      }
      return "mock-access-token";
    }

    if (this.pending) {
      const result = await this.pending.promise;
      return result.accessToken;
    }

    const account = await this.getCachedAccount();
    if (!account) {
      throw new Error(
        "Not logged in. Run `npm run login` or call the outlook_login tool, then complete device-code in the browser.",
      );
    }

    const silent = await this.pca!.acquireTokenSilent({
      account,
      scopes: this.config.scopes,
    });
    if (!silent?.accessToken) {
      throw new Error("Silent token refresh failed; run login again.");
    }
    this.persistAccount(silent.account);
    return silent.accessToken;
  }

  private async acquireByDeviceCode(
    onMessage: (message: string) => void,
  ): Promise<AuthenticationResult> {
    const result = await this.pca!.acquireTokenByDeviceCode({
      scopes: this.config.scopes,
      deviceCodeCallback: (response) => onMessage(response.message),
    });
    if (!result) {
      throw new Error("Device code login returned no result");
    }
    return result;
  }

  private async getCachedAccount(): Promise<AccountInfo | null> {
    const meta = loadSessionMeta(this.config.DATA_DIR);
    const accounts = await this.pca!.getTokenCache().getAllAccounts();
    if (meta?.homeAccountId) {
      const match = accounts.find((a) => a.homeAccountId === meta.homeAccountId);
      if (match) return match;
    }
    return accounts[0] ?? null;
  }

  private persistAccount(account: AccountInfo | null): void {
    if (!account) return;
    saveSessionMeta(this.config.DATA_DIR, {
      homeAccountId: account.homeAccountId,
      username: account.username,
      name: account.name,
      linkedAt: new Date().toISOString(),
    });
  }
}
