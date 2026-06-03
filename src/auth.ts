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

export async function getAccessToken(): Promise<TokenResult> {
  const config = loadConfig();
  const body = new URLSearchParams({
    grant_type: 'password',
    scope: config.scope,
    username: config.username,
    password: config.password,
    client_id: config.clientId,
  });

  const headers: Record<string, string> = {
    'Content-Type': 'application/x-www-form-urlencoded',
    ...config.headers,
  };

  const controller = new AbortController();
  const timeout = setTimeout(
    () => controller.abort(new Error(`Fetch timed out after ${config.timeoutMs}ms`)),
    config.timeoutMs,
  );

  try {
    const response = await fetch(new URL(config.tokenPath, config.baseUrl), {
      method: 'POST',
      headers,
      body: body.toString(),
      signal: controller.signal,
    });

    const text = await response.text();
    if (!response.ok) {
      throw new Error(`Token request failed with HTTP ${response.status}: ${text}`);
    }

    const raw = JSON.parse(text) as TokenResponse;
    if (!raw.access_token) {
      throw new Error('Token response did not include access_token');
    }

    return { token: raw.access_token, raw };
  } catch (error) {
    if (error instanceof Error && error.name === 'AbortError') {
      throw new Error(`Token request timed out after ${config.timeoutMs}ms`);
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}
