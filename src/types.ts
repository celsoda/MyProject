export type MailMessageSummary = {
  id: string;
  subject: string;
  from: string | null;
  receivedAt: string | null;
  preview: string | null;
  isRead: boolean;
};

export type AuthStatus = {
  linked: boolean;
  mode: "mock" | "live";
  username?: string;
  name?: string;
  homeAccountId?: string;
};

export type DeviceCodeStart = {
  userCode: string;
  verificationUri: string;
  message: string;
  expiresIn: number;
  interval: number;
};
