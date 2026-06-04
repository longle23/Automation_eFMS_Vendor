import dotenv from 'dotenv';

dotenv.config();

export interface AppConfig {
  baseUrl: string;
  tokenPath: string;
  clientId: string;
  scope: string;
  username: string;
  password: string;
  companyId?: string;
  authorization?: string;
  headers: Record<string, string>;
  timeoutMs: number;
  settlementPaymentUrl: string;
  settlementPaymentRequester: string;
  oneDrive?: OneDriveConfig;
}

export interface OneDriveConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  userId: string;
  fileId: string;
}

const DEFAULT_TOKEN_PATH = '/identityserver/connect/token';
const DEFAULT_SCOPE = 'openid profile offline_access efms_api';
const DEFAULT_CLIENT_ID = 'eFMS';
const DEFAULT_BASE_URL = 'https://efms-api.sotrans.com.vn';
const DEFAULT_SETTLEMENT_PAYMENT_URL =
  'https://efms-api.sotrans.com.vn/Accounting/api/v1/en-US/AcctSettlementPayment/paging?pageNumber=1&pageSize=1000';

export function loadConfig(): AppConfig {
  const baseUrl = process.env.EFMS_BASE_URL ?? DEFAULT_BASE_URL;
  const tokenPath = process.env.EFMS_TOKEN_PATH ?? DEFAULT_TOKEN_PATH;
  const clientId = process.env.EFMS_CLIENT_ID ?? DEFAULT_CLIENT_ID;
  const scope = process.env.EFMS_SCOPE ?? DEFAULT_SCOPE;
  const username = process.env.EFMS_USERNAME;
  const password = process.env.EFMS_PASSWORD;
  const companyId = process.env.EFMS_COMPANY_ID?.trim();
  const timeoutMs = Number(process.env.EFMS_TIMEOUT_MS ?? '15000');
  const settlementPaymentUrl = process.env.EFMS_SETTLEMENT_PAYMENT_URL ?? DEFAULT_SETTLEMENT_PAYMENT_URL;
  const settlementPaymentRequester = process.env.EFMS_SETTLEMENT_PAYMENT_REQUESTER?.trim();
  const oneDriveValues = {
    tenantId: process.env.ONEDRIVE_TENANT_ID?.trim(),
    clientId: process.env.ONEDRIVE_CLIENT_ID?.trim(),
    clientSecret: process.env.ONEDRIVE_CLIENT_SECRET?.trim(),
    userId: process.env.ONEDRIVE_USER_ID?.trim(),
    fileId: process.env.ONEDRIVE_FILE_ID?.trim(),
  };
  const configuredOneDriveValues = Object.values(oneDriveValues).filter(Boolean);
  const extraHeaders: Record<string, string> = {};

  if (companyId) {
    extraHeaders.companyId = companyId;
  }

  const authHeader = process.env.EFMS_AUTHORIZATION?.trim();
  if (authHeader) {
    extraHeaders.Authorization = authHeader;
  }

  if (!username || !password) {
    throw new Error('Missing EFMS_USERNAME or EFMS_PASSWORD in .env');
  }

  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
    throw new Error('EFMS_TIMEOUT_MS must be a positive number');
  }

  if (!settlementPaymentRequester) {
    throw new Error('Missing EFMS_SETTLEMENT_PAYMENT_REQUESTER in .env');
  }

  if (configuredOneDriveValues.length > 0 && configuredOneDriveValues.length < 5) {
    throw new Error('OneDrive configuration is incomplete in .env');
  }

  return {
    baseUrl,
    tokenPath,
    clientId,
    scope,
    username,
    password,
    companyId: companyId || undefined,
    authorization: authHeader || undefined,
    headers: extraHeaders,
    timeoutMs,
    settlementPaymentUrl,
    settlementPaymentRequester,
    oneDrive:
      configuredOneDriveValues.length === 5
        ? {
            tenantId: oneDriveValues.tenantId!,
            clientId: oneDriveValues.clientId!,
            clientSecret: oneDriveValues.clientSecret!,
            userId: oneDriveValues.userId!,
            fileId: oneDriveValues.fileId!,
          }
        : undefined,
  };
}
