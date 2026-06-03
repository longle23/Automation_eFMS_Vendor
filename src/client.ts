import { getAccessToken } from './auth.js';
import { loadConfig } from './config.js';

export interface ApiRequestOptions extends RequestInit {
  auth?: boolean;
}

export class EfmsClient {
  private readonly baseUrl: string;

  constructor(baseUrl = loadConfig().baseUrl) {
    this.baseUrl = baseUrl;
  }

  private buildUrl(pathname: string, query?: Record<string, string | number | boolean | undefined>): string {
    const url = new URL(pathname, this.baseUrl);

    if (query) {
      for (const [key, value] of Object.entries(query)) {
        if (value !== undefined && value !== null) {
          url.searchParams.set(key, String(value));
        }
      }
    }

    return url.toString();
  }

  async request(
    pathname: string,
    options: ApiRequestOptions = {},
    query?: Record<string, string | number | boolean | undefined>,
  ): Promise<Response> {
    const headers = new Headers(options.headers);

    if (options.auth !== false) {
      const { token } = await getAccessToken();
      headers.set('Authorization', `Bearer ${token}`);
    }

    return fetch(this.buildUrl(pathname, query), {
      ...options,
      headers,
    });
  }

  async getJson<T>(
    pathname: string,
    options: ApiRequestOptions = {},
    query?: Record<string, string | number | boolean | undefined>,
  ): Promise<T> {
    const response = await this.request(pathname, options, query);
    const text = await response.text();

    if (!response.ok) {
      throw new Error(`API request failed with HTTP ${response.status}: ${text}`);
    }

    return JSON.parse(text) as T;
  }

  async postJson<TBody extends object, TResponse>(
    pathname: string,
    body: TBody,
    options: ApiRequestOptions = {},
  ): Promise<TResponse> {
    const headers = new Headers(options.headers);
    headers.set('Content-Type', 'application/json');

    const response = await this.request(pathname, {
      ...options,
      auth: false,
      method: options.method ?? 'POST',
      headers,
      body: JSON.stringify(body),
    });
    const text = await response.text();

    if (!response.ok) {
      throw new Error(`API request failed with HTTP ${response.status}: ${text}`);
    }

    return JSON.parse(text) as TResponse;
  }
}
