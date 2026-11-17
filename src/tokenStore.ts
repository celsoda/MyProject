import { mkdirSync, readFileSync, writeFileSync, existsSync } from "node:fs";
import { join } from "node:path";
import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node";
import {
  decryptFromString,
  encryptToString,
  resolveEncryptionKey,
} from "./crypto.js";

export function createEncryptedCachePlugin(dataDir: string, keyMaterial?: string): ICachePlugin {
  mkdirSync(dataDir, { recursive: true });
  const cachePath = join(dataDir, "msal-cache.enc");
  const key = resolveEncryptionKey(dataDir, keyMaterial);

  return {
    async beforeCacheAccess(context: TokenCacheContext): Promise<void> {
      if (!existsSync(cachePath)) return;
      const encrypted = readFileSync(cachePath, "utf8");
      if (!encrypted.trim()) return;
      const json = decryptFromString(key, encrypted);
      context.tokenCache.deserialize(json);
    },
    async afterCacheAccess(context: TokenCacheContext): Promise<void> {
      if (!context.cacheHasChanged) return;
      const serialized = context.tokenCache.serialize();
      const encrypted = encryptToString(key, serialized);
      writeFileSync(cachePath, encrypted, { mode: 0o600 });
    },
  };
}

export type SessionMeta = {
  homeAccountId?: string;
  username?: string;
  name?: string;
  linkedAt?: string;
};

export function loadSessionMeta(dataDir: string): SessionMeta | null {
  const path = join(dataDir, "session.json");
  if (!existsSync(path)) return null;
  try {
    return JSON.parse(readFileSync(path, "utf8")) as SessionMeta;
  } catch {
    return null;
  }
}

export function saveSessionMeta(dataDir: string, meta: SessionMeta): void {
  mkdirSync(dataDir, { recursive: true });
  writeFileSync(join(dataDir, "session.json"), JSON.stringify(meta, null, 2), {
    mode: 0o600,
  });
}

export function clearSession(dataDir: string): void {
  for (const name of ["session.json", "msal-cache.enc"]) {
    const path = join(dataDir, name);
    if (existsSync(path)) {
      writeFileSync(path, "", { mode: 0o600 });
    }
  }
}
