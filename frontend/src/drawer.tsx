import type {
  KeyboardEvent as ReactKeyboardEvent,
  PointerEvent as ReactPointerEvent,
  ReactNode,
} from "react";

import type { PreviewRequest } from "./api";
import { React, ReactDOM, antd, currentTheme } from "./host";
import { t } from "./i18n";
import { OfficeViewer } from "./OfficeViewer";
import { installStyles } from "./styles";
import { usePreview } from "./usePreview";

const { ConfigProvider } = antd;

const WIDTH_STORAGE_KEY = "qwenpaw-office-preview-width";
const MIN_WIDTH = 440;
const MAX_WIDTH = 1280;

function clampWidth(value: number): number {
  const upper = Math.min(MAX_WIDTH, Math.max(MIN_WIDTH, window.innerWidth - 80));
  return Math.round(Math.min(upper, Math.max(MIN_WIDTH, value)));
}

function defaultWidth(): number {
  return clampWidth(Math.round(window.innerWidth * 0.46));
}

function readStoredWidth(): number {
  try {
    const stored = Number(window.localStorage.getItem(WIDTH_STORAGE_KEY));
    if (Number.isFinite(stored) && stored >= MIN_WIDTH) {
      return clampWidth(stored);
    }
  } catch {
    /* ignore */
  }
  return defaultWidth();
}

function useDrawerWidth() {
  const [width, setWidth] = React.useState(readStoredWidth);
  const drag = React.useRef<{ startX: number; startWidth: number } | null>(
    null,
  );

  React.useEffect(() => {
    const onResize = () => setWidth((current) => clampWidth(current));
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  const onPointerDown = (event: ReactPointerEvent<HTMLDivElement>) => {
    event.preventDefault();
    drag.current = { startX: event.clientX, startWidth: width };
    event.currentTarget.setPointerCapture?.(event.pointerId);
  };
  const onPointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (!drag.current) return;
    const delta = drag.current.startX - event.clientX;
    setWidth(clampWidth(drag.current.startWidth + delta));
  };
  const onPointerUp = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (!drag.current) return;
    drag.current = null;
    event.currentTarget.releasePointerCapture?.(event.pointerId);
    try {
      window.localStorage.setItem(WIDTH_STORAGE_KEY, String(width));
    } catch {
      /* ignore */
    }
  };
  const resetWidth = () => setWidth(defaultWidth());
  const nudgeWidth = (delta: number) =>
    setWidth((current) => {
      const next = clampWidth(current + delta);
      try {
        window.localStorage.setItem(WIDTH_STORAGE_KEY, String(next));
      } catch {
        /* ignore */
      }
      return next;
    });

  return {
    width,
    onPointerDown,
    onPointerMove,
    onPointerUp,
    resetWidth,
    nudgeWidth,
  };
}

interface DrawerHostProps {
  request: PreviewRequest;
  onClose: () => void;
}

function DrawerHost({ request, onClose }: DrawerHostProps) {
  const { state, run } = usePreview();
  const {
    width,
    onPointerDown,
    onPointerMove,
    onPointerUp,
    resetWidth,
    nudgeWidth,
  } = useDrawerWidth();
  const [closing, setClosing] = React.useState(false);
  const closingRef = React.useRef(false);
  const drawerRef = React.useRef<HTMLElement | null>(null);

  React.useEffect(() => {
    void run(request);
    drawerRef.current?.focus();
    // A new request remounts this component with a new ``key``.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const requestClose = React.useCallback(() => {
    if (closingRef.current) return;
    closingRef.current = true;
    setClosing(true);
    window.setTimeout(onClose, 180);
  }, [onClose]);

  React.useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.stopPropagation();
        requestClose();
      }
    };
    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, [requestClose]);

  const dark = currentTheme() === "dark";
  const algorithm =
    dark && antd.theme?.darkAlgorithm
      ? antd.theme.darkAlgorithm
      : antd.theme?.defaultAlgorithm;

  const contents = (
    <>
      <div
        className={`qopx-drawer-scrim${closing ? " qopx-closing" : ""}`}
        onMouseDown={() => requestClose()}
      />
      <aside
        ref={drawerRef}
        className={`qopx-drawer${closing ? " qopx-closing" : ""}`}
        style={{ width: `${width}px` }}
        role="dialog"
        aria-modal="true"
        aria-label={t("dialogLabel")}
        tabIndex={-1}
      >
        <div
          className="qopx-drawer-resizer"
          role="separator"
          aria-orientation="vertical"
          title={t("resizeHint")}
          onPointerDown={onPointerDown}
          onPointerMove={onPointerMove}
          onPointerUp={onPointerUp}
          onPointerCancel={onPointerUp}
          onDoubleClick={resetWidth}
          tabIndex={0}
          aria-label={t("resizeHint")}
          onKeyDown={(event: ReactKeyboardEvent) => {
            if (event.key === "ArrowLeft") {
              event.preventDefault();
              nudgeWidth(32);
            } else if (event.key === "ArrowRight") {
              event.preventDefault();
              nudgeWidth(-32);
            }
          }}
        />
        <OfficeViewer
          state={state}
          onClose={requestClose}
          onRetry={() => void run(request)}
          height="100%"
        />
      </aside>
    </>
  );

  if (!ConfigProvider) return contents as ReactNode;
  return (
    <ConfigProvider theme={algorithm ? { algorithm } : undefined}>
      {contents}
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
    container.id = "qwenpaw-office-preview-root";
    document.body.appendChild(container);
    root = ReactDOM.createRoot(container);
  }
  return root;
}

export function showOfficePreviewDrawer(request: PreviewRequest): void {
  installStyles();
  const reactRoot = ensureRoot();
  sequence += 1;
  if (!isOpen) {
    savedOverflow = document.body.style.overflow;
    isOpen = true;
  }
  document.body.style.overflow = "hidden";
  reactRoot.render(
    React.createElement(DrawerHost, {
      key: `preview-${sequence}`,
      request,
      onClose: closeOfficePreviewDrawer,
    }),
  );
}

export function closeOfficePreviewDrawer(): void {
  if (!isOpen) return;
  isOpen = false;
  if (root) {
    root.render(null);
  }
  document.body.style.overflow = savedOverflow;
  savedOverflow = "";
}
