import { readFile, writeFile } from 'node:fs/promises';

import type { OneDriveConfig } from './config.js';

type GraphTokenResponse = {
  access_token?: string;
  error?: string;
  error_description?: string;
};

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

export async function uploadFileToOneDrive(filePath: string, config: OneDriveConfig): Promise<void> {
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
    throw new Error(
      `OneDrive upload failed with HTTP ${uploadResponse.status}: ${await uploadResponse.text()}`,
    );
  }
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
    throw new Error(
      `OneDrive download failed with HTTP ${response.status}: ${await response.text()}`,
    );
  }

  await writeFile(filePath, Buffer.from(await response.arrayBuffer()));
}
