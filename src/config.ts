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
  api5Headers: Record<string, string>;
  timeoutMs: number;
  settlementPaymentUrl: string;
  settlementPaymentRequester: string;
  oneDrive?: OneDriveConfig;
}

export interface OneDriveConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  userId?: string;
  userPrincipalName?: string;
  driveId?: string;
  remotePath?: string;
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
    userPrincipalName: process.env.ONEDRIVE_USER_PRINCIPAL_NAME?.trim(),
    driveId: process.env.ONEDRIVE_DRIVE_ID?.trim(),
    remotePath: process.env.ONEDRIVE_REMOTE_PATH?.trim(),
    fileId: process.env.ONEDRIVE_FILE_ID?.trim(),
  };
  const hasOneDriveBaseConfig = Boolean(
      oneDriveValues.tenantId ||
      oneDriveValues.clientId ||
      oneDriveValues.clientSecret ||
      oneDriveValues.fileId ||
      oneDriveValues.userId ||
      oneDriveValues.userPrincipalName ||
      oneDriveValues.driveId,
  );
  const hasOneDriveConfig = Boolean(
      oneDriveValues.tenantId &&
      oneDriveValues.clientId &&
      oneDriveValues.clientSecret &&
      oneDriveValues.fileId &&
      (oneDriveValues.userId || oneDriveValues.userPrincipalName || oneDriveValues.driveId),
  );
  const extraHeaders: Record<string, string> = {};
  const api5Headers: Record<string, string> = {};

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

  if (hasOneDriveBaseConfig && !hasOneDriveConfig) {
    throw new Error('OneDrive configuration is incomplete in .env. Provide ONEDRIVE_FILE_ID and one of ONEDRIVE_USER_ID, ONEDRIVE_USER_PRINCIPAL_NAME, or ONEDRIVE_DRIVE_ID.');
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
    api5Headers,
    timeoutMs,
    settlementPaymentUrl,
    settlementPaymentRequester,
    oneDrive:
      hasOneDriveConfig
        ? {
            tenantId: oneDriveValues.tenantId!,
            clientId: oneDriveValues.clientId!,
            clientSecret: oneDriveValues.clientSecret!,
            userId: oneDriveValues.userId || undefined,
            userPrincipalName: oneDriveValues.userPrincipalName || undefined,
            driveId: oneDriveValues.driveId || undefined,
            remotePath: oneDriveValues.remotePath || undefined,
            fileId: oneDriveValues.fileId!,
          }
        : undefined,
  };
}
