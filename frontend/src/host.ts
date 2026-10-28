import type * as ReactNS from "react";

/**
 * The small slice of `window.QwenPaw.host` this plugin uses.  React and antd
 * are provided by the Console (never bundled), which keeps the plugin bundle
 * tiny and version-consistent with the host UI.
 */
export interface HostExternals {
  React: typeof ReactNS;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ReactDOM: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  antd: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  antdIcons: any;
  apiBaseUrl: string;
  getApiUrl(path: string): string;
  getApiToken(): string | null;
  useTheme?: () => "light" | "dark";
  useLocale?: () => string;
  getSelectedAgentId?: () => string;
  getCurrentSessionId?: () => string | null;
  openExternalLink?: (url: string) => void;
  fetch?: (path: string, init?: RequestInit) => Promise<Response>;
}

export interface PluginRouteDeclaration {
  path: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  component: React.ComponentType<any>;
  label: string | ((locale: string) => string);
  title?: string | ((locale: string) => string);
  icon?: string;
  priority?: number;
}

export interface QwenPawGlobal {
  host: HostExternals;
  registerRoutes?: (pluginId: string, routes: PluginRouteDeclaration[]) => void;
  modules?: Record<string, Record<string, unknown>>;
}

declare global {
  interface Window {
    QwenPaw: QwenPawGlobal;
  }
}

export const host: HostExternals = window.QwenPaw.host;
export const React: typeof ReactNS = host.React;
export const ReactDOM = host.ReactDOM;
export const antd = host.antd;
export const antdIcons = host.antdIcons;

export function currentTheme(): "light" | "dark" {
  try {
    if (document.documentElement.classList.contains("dark-mode")) {
      return "dark";
    }
    if (host.useTheme) {
      // Hooks cannot be called here; the class check above is the reliable
      // signal used by the Console itself.
    }
  } catch {
    /* ignore */
  }
  return "light";
}

export function readSelectedAgentId(): string {
  try {
    const raw =
      window.sessionStorage.getItem("qwenpaw-agent-storage") ??
      window.localStorage.getItem("qwenpaw-agent-storage");
    if (raw) {
      const parsed = JSON.parse(raw);
      const selected = parsed?.state?.selectedAgent;
      if (typeof selected === "string" && selected) return selected;
    }
  } catch {
    /* ignore */
  }
  try {
    return host.getSelectedAgentId?.() ?? "default";
  } catch {
    return "default";
  }
}

export function authHeaders(): Record<string, string> {
  const headers: Record<string, string> = {};
  try {
    const token = host.getApiToken();
    if (token) headers.Authorization = `Bearer ${token}`;
  } catch {
    /* ignore */
  }
  const agent = readSelectedAgentId();
  if (agent) headers["X-Agent-Id"] = agent;
  return headers;
}

/** Authenticated JSON/API request; falls back to manual headers on old hosts. */
export async function apiFetch(
  path: string,
  init: RequestInit = {},
): Promise<Response> {
  if (host.fetch) {
    return host.fetch(path, init);
  }
  const url = /^https?:\/\//i.test(path) ? path : host.getApiUrl(path);
  return fetch(url, {
    ...init,
    headers: {
      ...authHeaders(),
      ...((init.headers as Record<string, string>) ?? {}),
    },
  });
}

/** Fetch a raw URL (artifact preview URL, object URL, …). */
export async function fetchRaw(
  url: string,
  init: RequestInit = {},
): Promise<Response> {
  return fetch(url, {
    ...init,
    headers: {
      ...authHeaders(),
      ...((init.headers as Record<string, string>) ?? {}),
    },
  });
}
