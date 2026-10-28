import { baseName, isOfficeFile, type PreviewRequest } from "./api";
import { host } from "./host";
import { showOfficePreviewDrawer } from "./drawer";

const EVENT_NAME = "qwenpaw:open-file-preview";

interface FileTargetLike {
  source?: string;
  path?: string;
  root?: string;
  artifactUrl?: string;
}

interface PreviewEventDetail {
  target?: FileTargetLike;
  trigger?: HTMLElement | null;
}

function encodePath(path: string): string {
  return path
    .replace(/^\/+/, "")
    .split("/")
    .map((part) => encodeURIComponent(part))
    .join("/");
}

function buildRequest(target: FileTargetLike): PreviewRequest | null {
  const path = (target.path || "").trim();
  const artifactUrl = (target.artifactUrl || "").trim();
  if (!path && !artifactUrl) return null;
  if (!isOfficeFile(path || artifactUrl)) return null;

  const filename = baseName(path || artifactUrl);

  if (target.source === "workspace" && path) {
    return {
      filename,
      source: {
        kind: "workspace",
        path,
        root: target.root || "project",
      },
    };
  }
  if (artifactUrl) {
    return {
      filename,
      source: { kind: "artifact", url: artifactUrl, path },
    };
  }
  if (path.startsWith("/") || /^[a-zA-Z]:[\\/]/.test(path)) {
    // Historical chat artifact / absolute path without an explicit preview
    // URL. The host endpoint accepts both POSIX and ``C:/...`` paths.
    return {
      filename,
      source: {
        kind: "artifact",
        url: host.getApiUrl(`/files/preview${encodePath(path)}`),
        path,
      },
    };
  }
  if (path) {
    // A relative path from an "attachment" target is still a workspace path.
    return {
      filename,
      source: { kind: "workspace", path, root: target.root || "project" },
    };
  }
  return null;
}

function handlePreviewEvent(event: Event): void {
  const detail = (event as CustomEvent<PreviewEventDetail>).detail;
  const target = detail?.target;
  if (!target) return;
  const request = buildRequest(target);
  if (!request) return;

  // Office documents are not previewable by the host; stop here so the
  // built-in "preview unavailable" drawer does not fight with our modal.
  event.preventDefault();
  event.stopImmediatePropagation();
  showOfficePreviewDrawer(request);
}

let installed = false;

/** Hook into every internal file link the Console renders (chat cards,
 * artifacts, markdown links) without forking the host UI. */
export function installPreviewInterceptor(): void {
  if (installed) return;
  installed = true;
  window.addEventListener(EVENT_NAME, handlePreviewEvent, true);
}
