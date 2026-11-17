import { z } from "zod";

const envSchema = z.object({
  PORT: z.coerce.number().default(8787),
  /** Azure Entra application (client) ID — public client, no secret. */
  AZURE_CLIENT_ID: z.string().optional(),
  /**
   * Tenant: `common` (work+MSA), `organizations`, `consumers` (MSA only),
   * or a specific tenant GUID.
   */
  AZURE_TENANT: z.string().default("common"),
  /** AES-256 key as 64 hex chars, or any passphrase (hashed). Auto-generated for local demo if unset. */
  TOKEN_ENCRYPTION_KEY: z.string().optional(),
  /** Directory for encrypted MSAL cache (default .data). */
  DATA_DIR: z.string().default(".data"),
  /**
   * When `1`/`true`, skip Microsoft and serve mock inbox + fake auth.
   * Used for local smoke without Azure credentials.
   */
  MOCK_GRAPH: z
    .string()
    .optional()
    .transform((v) => v === "1" || v?.toLowerCase() === "true"),
  CONNECTOR_API_KEY: z.string().optional(),
});

export type AppConfig = z.infer<typeof envSchema> & {
  scopes: string[];
  authority: string;
  mockGraph: boolean;
  clientId: string;
};

/** Delegated user-consent scopes only (no application permissions). */
const DEFAULT_SCOPES = [
  "User.Read",
  "Mail.Read",
  "Mail.Send",
  "offline_access",
] as const;

export function loadConfig(env: NodeJS.ProcessEnv = process.env): AppConfig {
  const parsed = envSchema.parse(env);
  const mockGraph = Boolean(parsed.MOCK_GRAPH);
  const clientId = parsed.AZURE_CLIENT_ID ?? "";

  if (!mockGraph && !clientId) {
    throw new Error(
      "AZURE_CLIENT_ID is required unless MOCK_GRAPH=1. See README for Entra app registration.",
    );
  }

  return {
    ...parsed,
    mockGraph,
    clientId: clientId || "00000000-0000-0000-0000-000000000000",
    scopes: [...DEFAULT_SCOPES],
    authority: `https://login.microsoftonline.com/${parsed.AZURE_TENANT}`,
  };
}
