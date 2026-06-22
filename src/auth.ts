import { loadConfig } from './config.js';

export interface TokenResponse {
  access_token: string;
  expires_in?: number;
  token_type?: string;
  refresh_token?: string;
  scope?: string;
  [key: string]: unknown;
}

export interface TokenResult {
  token: string;
  raw: TokenResponse;
}

function buildTokenHeaders(config: ReturnType<typeof loadConfig>) {
  return {
    'Content-Type': 'application/x-www-form-urlencoded',
    ...config.headers,
  };
}

function buildApi5Headers(config: ReturnType<typeof loadConfig>) {
  return {
    ...buildTokenHeaders(config),
    departmentid: process.env.EFMS_API5_DEPARTMENT_ID?.trim() || '122',
    groupid: process.env.EFMS_API5_GROUP_ID?.trim() || '191',
    officeid: process.env.EFMS_API5_OFFICE_ID?.trim() || 'a38df6e1-5ecf-48af-91d0-be9ef32dcbf1',
  };
}

function buildApi9Headers(config: ReturnType<typeof loadConfig>) {
  return {
    'Content-Type': 'application/x-www-form-urlencoded',
    companyid: config.companyId || '',
    officeid: process.env.EFMS_API9_OFFICE_ID?.trim() || 'a1c27525-be96-48f4-91af-6ac7db6f6f92',
  };
}

async function requestToken(
  headers: Record<string, string>,
  endpoint = '/identityserver/connect/token',
  bodyOverrides: Partial<Record<'grant_type' | 'scope' | 'username' | 'password' | 'client_id', string>> = {},
): Promise<{ response: Response; text: string; headers: Record<string, string> }> {
  const config = loadConfig();
  const body = new URLSearchParams({
    grant_type: bodyOverrides.grant_type ?? 'password',
    scope: bodyOverrides.scope ?? config.scope,
    username: bodyOverrides.username ?? config.username,
    password: bodyOverrides.password ?? config.password,
    client_id: bodyOverrides.client_id ?? config.clientId,
  });

  const controller = new AbortController();
  const timeout = setTimeout(
    () => controller.abort(new Error(`Fetch timed out after ${config.timeoutMs}ms`)),
    config.timeoutMs,
  );

  try {
    const response = await fetch(new URL(endpoint, config.baseUrl), {
      method: 'POST',
      headers,
      body: body.toString(),
      signal: controller.signal,
    });
    const text = await response.text();
    return { response, text, headers };
  } finally {
    clearTimeout(timeout);
  }
}

export async function getAccessToken(): Promise<TokenResult> {
  const config = loadConfig();
  const { response, text } = await requestToken(buildTokenHeaders(config));

  if (!response.ok) {
    throw new Error(`Token request failed with HTTP ${response.status}: ${text}`);
  }

  const raw = JSON.parse(text) as TokenResponse;
  if (!raw.access_token) {
    throw new Error('Token response did not include access_token');
  }

  return { token: raw.access_token, raw };
}

export async function getApi5TokenResponse() {
  const config = loadConfig();
  const { response, text, headers } = await requestToken(buildApi5Headers(config));
  return { response, text, headers };
}

export async function getApi9TokenResponse() {
  const config = loadConfig();
  const { response, text, headers } = await requestToken(
    buildApi9Headers(config),
    '/identityserver/connect/token',
    {
      grant_type: 'password',
      scope: config.scope,
      username: config.username,
      password: config.password,
      client_id: config.clientId,
    },
  );
  return { response, text, headers };
}
