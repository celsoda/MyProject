import type { MailMessageSummary } from "./types.js";

type GraphMessage = {
  id: string;
  subject?: string;
  bodyPreview?: string;
  receivedDateTime?: string;
  isRead?: boolean;
  from?: { emailAddress?: { name?: string; address?: string } };
};

function mapMessage(m: GraphMessage): MailMessageSummary {
  const fromAddr = m.from?.emailAddress;
  const from = fromAddr
    ? [fromAddr.name, fromAddr.address].filter(Boolean).join(" <") +
      (fromAddr.name && fromAddr.address ? ">" : "")
    : null;
  return {
    id: m.id,
    subject: m.subject ?? "(no subject)",
    from,
    receivedAt: m.receivedDateTime ?? null,
    preview: m.bodyPreview ?? null,
    isRead: Boolean(m.isRead),
  };
}

export async function listInboxMessages(
  accessToken: string,
  limit = 10,
): Promise<MailMessageSummary[]> {
  const top = Math.min(Math.max(limit, 1), 50);
  const url = new URL(
    "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages",
  );
  url.searchParams.set("$top", String(top));
  url.searchParams.set("$orderby", "receivedDateTime desc");
  url.searchParams.set(
    "$select",
    "id,subject,from,receivedDateTime,bodyPreview,isRead",
  );

  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
    },
  });

  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph list messages failed (${res.status}): ${body}`);
  }

  const json = (await res.json()) as { value?: GraphMessage[] };
  return (json.value ?? []).map(mapMessage);
}

export async function sendMail(
  accessToken: string,
  input: { to: string[]; subject: string; bodyText: string },
): Promise<{ ok: true }> {
  const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message: {
        subject: input.subject,
        body: { contentType: "Text", content: input.bodyText },
        toRecipients: input.to.map((address) => ({
          emailAddress: { address },
        })),
      },
      saveToSentItems: true,
    }),
  });

  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph sendMail failed (${res.status}): ${body}`);
  }

  return { ok: true };
}

export const MOCK_MESSAGES: MailMessageSummary[] = [
  {
    id: "mock-1",
    subject: "Welcome to the Outlook connector spike",
    from: "Connector <noreply@example.com>",
    receivedAt: new Date().toISOString(),
    preview: "Mock inbox message for local smoke tests (MOCK_GRAPH=1).",
    isRead: false,
  },
  {
    id: "mock-2",
    subject: "Second mock message",
    from: "Demo <demo@example.com>",
    receivedAt: new Date(Date.now() - 3600_000).toISOString(),
    preview: "Claude Desktop can call list_mail without Microsoft tokens.",
    isRead: true,
  },
];
