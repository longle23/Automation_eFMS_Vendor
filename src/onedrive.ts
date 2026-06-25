import { readFile, writeFile } from 'node:fs/promises';

import type { OneDriveConfig } from './config.js';

type GraphTokenResponse = {
  access_token?: string;
  error?: string;
  error_description?: string;
};

type DriveItemResponse = {
  id?: string;
  name?: string;
};

export class OneDriveRequestError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly responseText: string,
  ) {
    super(message);
  }
}

async function getGraphAccessToken(config: OneDriveConfig): Promise<string> {
  const response = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: new URLSearchParams({
        client_id: config.clientId,
        client_secret: config.clientSecret,
        grant_type: 'client_credentials',
        scope: 'https://graph.microsoft.com/.default',
      }),
    },
  );
  const body = (await response.json()) as GraphTokenResponse;

  if (!response.ok || !body.access_token) {
    throw new Error(
      `OneDrive token request failed with HTTP ${response.status}: ${
        body.error_description ?? body.error ?? 'missing access_token'
      }`,
    );
  }

  return body.access_token;
}

function buildDriveItemUrl(config: OneDriveConfig, suffix = ''): string {
  const encodedFileId = encodeURIComponent(config.fileId);

  if (config.driveId) {
    return `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(config.driveId)}/items/${encodedFileId}${suffix}`;
  }

  if (config.userId) {
    return `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(config.userId)}/drive/items/${encodedFileId}${suffix}`;
  }

  if (config.userPrincipalName) {
    return `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(config.userPrincipalName)}/drive/items/${encodedFileId}${suffix}`;
  }

  throw new Error('OneDrive configuration requires ONEDRIVE_USER_ID, ONEDRIVE_USER_PRINCIPAL_NAME, or ONEDRIVE_DRIVE_ID');
}

function encodeDrivePath(path: string): string {
  return path
    .split(/[\\/]+/)
    .filter(Boolean)
    .map(encodeURIComponent)
    .join('/');
}

function buildDrivePathContentUrl(config: OneDriveConfig, remoteFileName: string): string {
  if (!config.remotePath) {
    throw new Error('ONEDRIVE_REMOTE_PATH is required to upload a missing workbook by path');
  }

  const remotePath = encodeDrivePath(`${config.remotePath}/${remoteFileName}`);

  if (config.driveId) {
    return `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(config.driveId)}/root:/${remotePath}:/content`;
  }

  if (config.userId) {
    return `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(config.userId)}/drive/root:/${remotePath}:/content`;
  }

  if (config.userPrincipalName) {
    return `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(config.userPrincipalName)}/drive/root:/${remotePath}:/content`;
  }

  throw new Error('OneDrive configuration requires ONEDRIVE_USER_ID, ONEDRIVE_USER_PRINCIPAL_NAME, or ONEDRIVE_DRIVE_ID');
}

export function isOneDriveNotFoundError(error: unknown): boolean {
  return (
    error instanceof OneDriveRequestError &&
    (error.status === 404 || error.responseText.includes('itemNotFound'))
  );
}

export async function checkOneDriveFile(config: OneDriveConfig): Promise<string> {
  const token = await getGraphAccessToken(config);
  const response = await fetch(
    buildDriveItemUrl(config),
    {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    },
  );

  if (!response.ok) {
    throw new Error(`OneDrive file check failed with HTTP ${response.status}: ${await response.text()}`);
  }

  const item = (await response.json()) as { name?: string; file?: unknown };
  if (!item.file) {
    throw new Error(`OneDrive item is not a file: ${config.fileId}`);
  }

  return item.name ?? config.fileId;
}

export async function uploadFileToOneDrive(
  filePath: string,
  config: OneDriveConfig,
): Promise<DriveItemResponse> {
  const token = await getGraphAccessToken(config);
  const file = await readFile(filePath);
  const uploadResponse = await fetch(
    buildDriveItemUrl(config, '/content'),
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      body: file,
    },
  );

  if (!uploadResponse.ok) {
    const responseText = await uploadResponse.text();
    throw new OneDriveRequestError(
      `OneDrive upload failed with HTTP ${uploadResponse.status}: ${responseText}`,
      uploadResponse.status,
      responseText,
    );
  }

  return (await uploadResponse.json()) as DriveItemResponse;
}

export async function uploadFileToOneDrivePath(
  filePath: string,
  config: OneDriveConfig,
  remoteFileName: string,
): Promise<DriveItemResponse> {
  const token = await getGraphAccessToken(config);
  const file = await readFile(filePath);
  const uploadResponse = await fetch(
    buildDrivePathContentUrl(config, remoteFileName),
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      body: file,
    },
  );

  if (!uploadResponse.ok) {
    const responseText = await uploadResponse.text();
    throw new OneDriveRequestError(
      `OneDrive path upload failed with HTTP ${uploadResponse.status}: ${responseText}`,
      uploadResponse.status,
      responseText,
    );
  }

  return (await uploadResponse.json()) as DriveItemResponse;
}

export async function downloadFileFromOneDrive(filePath: string, config: OneDriveConfig): Promise<void> {
  const token = await getGraphAccessToken(config);
  const response = await fetch(
    buildDriveItemUrl(config, '/content'),
    {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    },
  );

  if (!response.ok) {
    const responseText = await response.text();
    throw new OneDriveRequestError(
      `OneDrive download failed with HTTP ${response.status}: ${responseText}`,
      response.status,
      responseText,
    );
  }

  await writeFile(filePath, Buffer.from(await response.arrayBuffer()));
}
