import { Hono } from "hono";
import { HTTPException } from "hono/http-exception";
import type { Connector } from "./connector.js";

function requireApiKey(
  apiKey: string | undefined,
  header: string | undefined,
): void {
  if (!apiKey) return;
  if (header !== apiKey) {
    throw new HTTPException(401, { message: "Invalid or missing x-api-key" });
  }
}

export function createHttpApp(connector: Connector, apiKey?: string): Hono {
  const app = new Hono();

  app.onError((err, c) => {
    if (err instanceof HTTPException) {
      return err.getResponse();
    }
    const message = err instanceof Error ? err.message : "Unknown error";
    const status = /not logged in/i.test(message) ? 401 : 500;
    return c.json({ error: message }, status);
  });

  app.get("/health", (c) => c.json({ ok: true, service: "outlook-connector" }));

  app.use("/v1/*", async (c, next) => {
    requireApiKey(apiKey, c.req.header("x-api-key") ?? undefined);
    await next();
  });

  app.get("/v1/auth/status", (c) => c.json(connector.authStatus()));

  app.post("/v1/auth/device-code/start", async (c) => {
    const info = await connector.startLogin();
    return c.json(info);
  });

  app.post("/v1/auth/device-code/wait", async (c) => {
    const status = await connector.waitLogin();
    return c.json(status);
  });

  app.delete("/v1/auth", async (c) => {
    await connector.logout();
    return c.json({ ok: true });
  });

  app.get("/v1/mail/messages", async (c) => {
    const limit = Number(c.req.query("limit") ?? "10");
    const messages = await connector.listMail(Number.isFinite(limit) ? limit : 10);
    return c.json({ messages });
  });

  app.post("/v1/mail/messages", async (c) => {
    const body = await c.req.json<{
      to?: string[];
      subject?: string;
      bodyText?: string;
    }>();
    if (!body.to?.length || !body.subject || body.bodyText == null) {
      throw new HTTPException(400, {
        message: "Body must include to[], subject, bodyText",
      });
    }
    const result = await connector.sendMail({
      to: body.to,
      subject: body.subject,
      bodyText: body.bodyText,
    });
    return c.json(result);
  });

  return app;
}
