import { apiFetch, currentTheme, fetchRaw } from "./host";

export const OFFICE_EXTENSIONS = new Set([
  "csv",
  "doc",
  "docm",
  "docx",
  "dot",
  "dotm",
  "dotx",
  "dps",
  "et",
  "pps",
  "ppsx",
  "ppt",
  "pptm",
  "pptx",
  "potx",
  "tsv",
  "wps",
  "xls",
  "xlsm",
  "xlsx",
  "xltm",
  "xltx",
]);

export const ACCEPT_ATTR = Array.from(OFFICE_EXTENSIONS)
  .map((ext) => `.${ext}`)
  .join(",");

export interface ConvertResult {
  ok: boolean;
  filename: string;
  kind: "docx" | "pptx" | "xlsx" | "csv" | string;
  format: "html" | string;
  html: string;
  warnings: string[];
  meta: {
    kind: string;
    label: string;
    extension: string;
    converted_from: string | null;
    size: number;
    html_size: number;
  };
  cached?: boolean;
}

export interface HealthResult {
  ok: boolean;
  supported_extensions: string[];
  renderers: Record<string, boolean>;
  libreoffice: string | null;
  max_mb: number;
  notes: string[];
}

export interface DirectoryEntry {
  name: string;
  path: string;
  kind: "file" | "directory";
  size: number | null;
  modified_at: string;
  preview_kind: string;
}

export interface DirectoryPage {
  directory: string;
  entries: DirectoryEntry[];
  has_more: boolean;
  next_cursor: string | null;
}

export type PreviewSource =
  | { kind: "upload"; blob: Blob }
  | { kind: "workspace"; path: string; root?: string }
  | { kind: "artifact"; url: string; path?: string };

export interface PreviewRequest {
  filename: string;
  source: PreviewSource;
}

export class ApiError extends Error {
  status: number;
  code: string;

  constructor(message: string, status: number, code: string) {
    super(message);
    this.name = "ApiError";
    this.status = status;
    this.code = code;
  }
}

export function extensionOf(path: string): string {
  const clean = (path || "").split(/[?#]/, 1)[0];
  const name = clean.replace(/\\/g, "/").split("/").pop() || "";
  const dot = name.lastIndexOf(".");
  return dot >= 0 ? name.slice(dot + 1).toLowerCase() : "";
}

export function isOfficeFile(path: string): boolean {
  return OFFICE_EXTENSIONS.has(extensionOf(path));
}

export function baseName(path: string): string {
  const clean = (path || "").replace(/\\/g, "/").split(/[?#]/, 1)[0];
  return clean.split("/").filter(Boolean).pop() || "document";
}

export async function readApiError(response: Response): Promise<ApiError> {
  const status = response.status;
  let message = `${status} ${response.statusText}`.trim();
  let code = `http_${status}`;
  try {
    const data = await response.json();
    const detail = data?.detail ?? data;
    if (typeof detail === "string") {
      message = detail;
    } else if (detail && typeof detail === "object") {
      if (typeof detail.message === "string") message = detail.message;
      if (typeof detail.code === "string") code = detail.code;
    }
  } catch {
    try {
      const text = await response.text();
      if (text) message = text.slice(0, 500);
    } catch {
      /* ignore */
    }
  }
  return new ApiError(message, status, code);
}

export async function fetchHealth(): Promise<HealthResult> {
  const response = await apiFetch("/office-preview/health");
  if (!response.ok) throw await readApiError(response);
  return (await response.json()) as HealthResult;
}

export async function listWorkspace(
  path: string,
  root: string,
  limit = 400,
): Promise<DirectoryPage> {
  const params = new URLSearchParams({
    path,
    root,
    limit: String(limit),
  });
  const response = await apiFetch(`/workspace/tree?${params.toString()}`);
  if (!response.ok) throw await readApiError(response);
  return (await response.json()) as DirectoryPage;
}

export interface ResolvedSource {
  blob: Blob;
  filename: string;
}

export async function resolveSource(
  request: PreviewRequest,
): Promise<ResolvedSource> {
  const { source } = request;
  if (source.kind === "upload") {
    return { blob: source.blob, filename: request.filename };
  }

  let response: Response;
  if (source.kind === "workspace") {
    const params = new URLSearchParams({
      path: source.path,
      root: source.root || "project",
    });
    response = await apiFetch(`/workspace/file-download?${params.toString()}`);
  } else {
    response = await fetchRaw(source.url);
  }
  if (!response.ok) throw await readApiError(response);
  const blob = await response.blob();
  const filename =
    request.filename ||
    baseName(source.kind === "workspace" ? source.path : source.path || "");
  return { blob, filename };
}

export async function convertBlob(
  blob: Blob,
  filename: string,
  signal?: AbortSignal,
): Promise<ConvertResult> {
  const params = new URLSearchParams({
    filename: filename || "document.docx",
    theme: currentTheme(),
  });
  const response = await apiFetch(`/office-preview/convert?${params.toString()}`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream" },
    body: blob,
    signal,
  });
  if (!response.ok) throw await readApiError(response);
  const result = (await response.json()) as ConvertResult;
  if (!result || typeof result.html !== "string" || !result.html) {
    throw new ApiError(
      "The server returned an empty preview.",
      500,
      "empty_result",
    );
  }
  return result;
}
