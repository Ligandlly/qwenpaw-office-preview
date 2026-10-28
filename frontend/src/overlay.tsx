import type { KeyboardEvent as ReactKeyboardEvent, MouseEvent as ReactMouseEvent, ReactNode } from "react";

import type { PreviewRequest } from "./api";
import {
  React,
  ReactDOM,
  antd,
  currentTheme,
} from "./host";
import { t } from "./i18n";
import { OfficeViewer } from "./OfficeViewer";
import { installStyles } from "./styles";
import { usePreview } from "./usePreview";

const { ConfigProvider } = antd;

interface OverlayHostProps {
  request: PreviewRequest;
  onClose: () => void;
}

function OverlayHost({ request, onClose }: OverlayHostProps) {
  const { state, run } = usePreview();

  React.useEffect(() => {
    void run(request);
    // A new request is delivered by remounting this component with a new
    // ``key``; run identity is stable.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  React.useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.stopPropagation();
        onClose();
      }
    };
    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, [onClose]);

  const dark = currentTheme() === "dark";
  const algorithm =
    dark && antd.theme?.darkAlgorithm
      ? antd.theme.darkAlgorithm
      : antd.theme?.defaultAlgorithm;

  const panel = (
    <div
      className="qopx-overlay"
      onMouseDown={(event: ReactMouseEvent<HTMLDivElement>) => {
        if (event.target === event.currentTarget) onClose();
      }}
    >
      <div
        className="qopx-overlay-panel"
        role="dialog"
        aria-modal="true"
        aria-label={t("dialogLabel")}
      >
        <OfficeViewer
          state={state}
          onClose={onClose}
          onRetry={() => void run(request)}
          height="100%"
        />
      </div>
    </div>
  );

  if (!ConfigProvider) return panel as ReactNode;
  return (
    <ConfigProvider theme={algorithm ? { algorithm } : undefined}>
      {panel}
    </ConfigProvider>
  );
}

let container: HTMLDivElement | null = null;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
let root: any = null;
let sequence = 0;
let savedOverflow = "";
let isOpen = false;

function ensureRoot() {
  if (!container) {
    container = document.createElement("div");
    container.id = "qwenpaw-office-preview-overlay";
    document.body.appendChild(container);
    root = ReactDOM.createRoot(container);
  }
  return root;
}

export function showOfficePreviewOverlay(request: PreviewRequest): void {
  installStyles();
  const reactRoot = ensureRoot();
  sequence += 1;
  if (!isOpen) {
    savedOverflow = document.body.style.overflow;
    isOpen = true;
  }
  document.body.style.overflow = "hidden";
  reactRoot.render(
    React.createElement(OverlayHost, {
      key: `preview-${sequence}`,
      request,
      onClose: closeOfficePreviewOverlay,
    }),
  );
}

export function closeOfficePreviewOverlay(): void {
  if (!isOpen) return;
  isOpen = false;
  if (root) {
    root.render(null);
  }
  document.body.style.overflow = savedOverflow;
  savedOverflow = "";
}
