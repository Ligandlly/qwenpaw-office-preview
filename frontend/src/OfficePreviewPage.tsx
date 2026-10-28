import type {
  ChangeEvent,
  DragEvent as ReactDragEvent,
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent as ReactMouseEvent,
  ReactNode,
} from "react";

import {
  ACCEPT_ATTR,
  baseName,
  fetchHealth,
  isOfficeFile,
  listWorkspace,
  type DirectoryEntry,
  type DirectoryPage,
  type HealthResult,
  type PreviewRequest,
} from "./api";
import { React, antd, antdIcons, currentTheme } from "./host";
import { formatBytes, formatDateTime, fileEmoji } from "./format";
import { t } from "./i18n";
import { OfficeViewer } from "./OfficeViewer";
import { usePreview } from "./usePreview";

const { Alert, Button, Empty, Radio, Space, Spin, Switch, Tag, Typography } = antd;
const { ReloadOutlined, UploadOutlined } = antdIcons;

function readQueryParams(): URLSearchParams {
  const hash = window.location.hash || "";
  const marker = hash.indexOf("?");
  if (marker >= 0) {
    return new URLSearchParams(hash.slice(marker + 1));
  }
  return new URLSearchParams(window.location.search);
}

function HealthChips() {
  const [health, setHealth] = React.useState<HealthResult | null>(null);
  const [failed, setFailed] = React.useState(false);

  React.useEffect(() => {
    let alive = true;
    fetchHealth()
      .then((value) => {
        if (alive) setHealth(value);
      })
      .catch(() => {
        if (alive) setFailed(true);
      });
    return () => {
      alive = false;
    };
  }, []);

  if (failed) {
    return (
      <Typography.Text type="secondary" style={{ fontSize: 12 }}>
        {t("healthFailed")}
      </Typography.Text>
    );
  }
  if (!health) return null;

  const entries: Array<{ label: string; ok: boolean }> = [
    { label: "Word", ok: Boolean(health.renderers.docx) },
    { label: "PowerPoint", ok: Boolean(health.renderers.pptx) },
    { label: "Excel", ok: Boolean(health.renderers.xlsx) },
    { label: "LibreOffice", ok: Boolean(health.libreoffice) },
  ];
  const missing = entries.some((entry) => !entry.ok && entry.label !== "LibreOffice");

  return (
    <>
      <div className="qopx-status">
        <span className="qopx-chip">
          {t("status")}
        </span>
        {entries.map((entry) => (
          <span className="qopx-chip" key={entry.label}>
            <span
              className={`qopx-dot ${entry.ok ? "" : "qopx-bad"}`}
              aria-hidden="true"
            />
            {entry.label}: {entry.ok ? t("statusOk") : t("statusMissing")}
          </span>
        ))}
      </div>
      {missing && (
        <Alert
          type="warning"
          showIcon
          style={{ marginTop: 12 }}
          message={t("depsMissing")}
          description={
            health.notes?.length ? (
              <ul className="qopx-warning-list">
                {health.notes.map((note) => (
                  <li key={note}>{note}</li>
                ))}
              </ul>
            ) : undefined
          }
        />
      )}
      {!health.libreoffice && (
        <Typography.Paragraph
          type="secondary"
          style={{ fontSize: 12, margin: "10px 0 0" }}
        >
          {t("legacyHint")}
        </Typography.Paragraph>
      )}
    </>
  );
}

interface DropZoneProps {
  onFile: (file: File) => void;
}

function DropZone({ onFile }: DropZoneProps) {
  const inputRef = React.useRef<HTMLInputElement | null>(null);
  const [dragging, setDragging] = React.useState(false);

  const accept = (files: FileList | null) => {
    const file = files && files[0];
    if (!file) return;
    if (!isOfficeFile(file.name)) {
      antd.message.warning(`${t("unsupported")} · ${t("unsupportedHint")}`);
      return;
    }
    onFile(file);
  };

  return (
    <div
      className={`qopx-drop${dragging ? " qopx-dragging" : ""}`}
      role="button"
      tabIndex={0}
      onClick={() => inputRef.current?.click()}
      onKeyDown={(event: ReactKeyboardEvent) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          inputRef.current?.click();
        }
      }}
      onDragOver={(event: ReactDragEvent) => {
        event.preventDefault();
        setDragging(true);
      }}
      onDragLeave={(event: ReactDragEvent) => {
        event.preventDefault();
        setDragging(false);
      }}
      onDrop={(event: ReactDragEvent) => {
        event.preventDefault();
        setDragging(false);
        accept(event.dataTransfer?.files ?? null);
      }}
    >
      <span className="qopx-drop-icon" aria-hidden="true">
        📄
      </span>
      <span className="qopx-drop-title">{t("dropTitle")}</span>
      <span className="qopx-drop-hint">{t("dropHint")}</span>
      <div className="qopx-drop-actions">
        <Button
          type="primary"
          size="small"
          icon={<UploadOutlined />}
          onClick={(event: ReactMouseEvent) => {
            event.stopPropagation();
            inputRef.current?.click();
          }}
        >
          {t("chooseFile")}
        </Button>
      </div>
      <input
        ref={inputRef}
        type="file"
        accept={ACCEPT_ATTR}
        style={{ display: "none" }}
        onChange={(event: ChangeEvent<HTMLInputElement>) => {
          accept(event.target.files);
          event.target.value = "";
        }}
      />
    </div>
  );
}

interface WorkspaceBrowserProps {
  onPick: (request: PreviewRequest) => void;
}

function WorkspaceBrowser({ onPick }: WorkspaceBrowserProps) {
  const [root, setRoot] = React.useState("project");
  const [path, setPath] = React.useState("");
  const [page, setPage] = React.useState<DirectoryPage | null>(null);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState("");
  const [onlyPreviewable, setOnlyPreviewable] = React.useState(true);
  const sequence = React.useRef(0);

  const load = React.useCallback(
    async (targetPath: string, targetRoot: string) => {
      const id = sequence.current + 1;
      sequence.current = id;
      setLoading(true);
      setError("");
      try {
        const data = await listWorkspace(targetPath, targetRoot);
        if (id === sequence.current) setPage(data);
      } catch (cause) {
        if (id === sequence.current) {
          setPage(null);
          setError((cause as Error)?.message || t("loadDirFailed"));
        }
      } finally {
        if (id === sequence.current) setLoading(false);
      }
    },
    [],
  );

  React.useEffect(() => {
    void load(path, root);
  }, [path, root, load]);

  const segments = path.split("/").filter(Boolean);
  const entries = (page?.entries ?? []).filter((entry) => {
    if (entry.kind === "directory") return true;
    return !onlyPreviewable || isOfficeFile(entry.path);
  });

  const goUp = () => setPath(segments.slice(0, -1).join("/"));

  return (
    <>
      <div className="qopx-breadcrumb">
        <Radio.Group
          size="small"
          value={root}
          optionType="button"
          options={[
            { label: t("projectRoot"), value: "project" },
            { label: t("workspaceRoot"), value: "workspace" },
          ]}
          onChange={(event: { target: { value: string } }) => {
            setPath("");
            setRoot(event.target.value);
          }}
        />
        <span style={{ width: 8 }} />
        <button
          type="button"
          className="qopx-crumb"
          onClick={() => setPath("")}
          title={t("root")}
        >
          {root === "workspace" ? t("workspaceRoot") : t("projectRoot")}
        </button>
        {segments.map((segment, index) => (
          <React.Fragment key={`${segment}-${index}`}>
            <span className="qopx-crumb-sep">/</span>
            <button
              type="button"
              className="qopx-crumb"
              onClick={() => setPath(segments.slice(0, index + 1).join("/"))}
            >
              {segment}
            </button>
          </React.Fragment>
        ))}
        <span style={{ flex: 1 }} />
        <Space size={6}>
          {path && (
            <Button size="small" type="text" onClick={goUp}>
              {t("parent")}
            </Button>
          )}
          <Button
            size="small"
            type="text"
            icon={<ReloadOutlined />}
            onClick={() => void load(path, root)}
          >
            {t("refresh")}
          </Button>
          <Typography.Text type="secondary" style={{ fontSize: 12 }}>
            {t("filterOnly")}
          </Typography.Text>
          <Switch
            size="small"
            checked={onlyPreviewable}
            onChange={setOnlyPreviewable}
          />
        </Space>
      </div>

      {error && (
        <div style={{ padding: 16 }}>
          <Alert type="error" showIcon message={t("loadDirFailed")} description={error} />
        </div>
      )}

      <div className="qopx-file-list">
        {loading ? (
          <div className="qopx-empty">
            <Spin size="small" /> <span style={{ marginLeft: 8 }}>{t("loading")}</span>
          </div>
        ) : entries.length === 0 ? (
          <div className="qopx-empty">
            <Empty
              image={Empty.PRESENTED_IMAGE_SIMPLE}
              description={t("emptyDir")}
            />
          </div>
        ) : (
          entries.map((entry: DirectoryEntry) => {
            const supported = entry.kind === "directory" || isOfficeFile(entry.path);
            return (
              <button
                type="button"
                key={entry.path}
                className="qopx-file-row"
                disabled={!supported}
                onClick={() => {
                  if (entry.kind === "directory") {
                    setPath(entry.path);
                    return;
                  }
                  onPick({
                    filename: entry.name,
                    source: { kind: "workspace", path: entry.path, root },
                  });
                }}
              >
                <span className="qopx-file-icon" aria-hidden="true">
                  {entry.kind === "directory" ? "📁" : fileEmoji(entry.name)}
                </span>
                <span className="qopx-file-main">
                  <span className="qopx-file-name" title={entry.name}>
                    {entry.name}
                  </span>
                  <span className="qopx-file-sub">
                    {entry.kind === "directory"
                      ? "folder"
                      : [
                          formatBytes(entry.size),
                          formatDateTime(entry.modified_at),
                        ]
                          .filter(Boolean)
                          .join(" · ")}
                  </span>
                </span>
                {entry.kind === "file" && !supported && (
                  <Tag>{t("unsupported")}</Tag>
                )}
              </button>
            );
          })
        )}
      </div>
    </>
  );
}

export function OfficePreviewPage() {
  const { state, run, reset } = usePreview();
  const [lastRequest, setLastRequest] = React.useState<PreviewRequest | null>(
    null,
  );
  const viewerRef = React.useRef<HTMLElement | null>(null);
  const initialised = React.useRef(false);

  const start = React.useCallback(
    (request: PreviewRequest) => {
      setLastRequest(request);
      void run(request);
      window.setTimeout(() => {
        viewerRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
      }, 30);
    },
    [run],
  );

  React.useEffect(() => {
    if (initialised.current) return;
    initialised.current = true;
    const params = readQueryParams();
    const artifact = params.get("artifact");
    const path = params.get("path");
    const name = params.get("name") || undefined;
    if (artifact) {
      start({
        filename: name || baseName(artifact),
        source: { kind: "artifact", url: artifact },
      });
    } else if (path) {
      start({
        filename: name || baseName(path),
        source: {
          kind: "workspace",
          path,
          root: params.get("root") || "project",
        },
      });
    }
  }, [start]);

  let viewer: ReactNode = null;
  if (state.status !== "idle") {
    viewer = (
      <section className="qopx-card" ref={viewerRef}>
        <OfficeViewer
          state={state}
          onClose={reset}
          onRetry={lastRequest ? () => void run(lastRequest) : undefined}
          height="70vh"
        />
      </section>
    );
  }

  return (
    <div className="qopx-page" data-theme={currentTheme()}>
      <header className="qopx-hero">
        <h1>
          <span className="qopx-hero-icon" aria-hidden="true">
            📄
          </span>
          {t("routeTitle")}
        </h1>
        <p>{t("pageSubtitle")}</p>
      </header>

      <div className="qopx-grid">
        <section className="qopx-card">
          <div className="qopx-card-body">
            <DropZone
              onFile={(file) =>
                start({
                  filename: file.name,
                  source: { kind: "upload", blob: file },
                })
              }
            />
            <HealthChips />
          </div>
        </section>

        {viewer}

        <section className="qopx-card">
          <div className="qopx-card-head">
            <div className="qopx-card-title">🗂 {t("workspace")}</div>
            <Typography.Text type="secondary" style={{ fontSize: 12 }}>
              {t("cacheNote")}
            </Typography.Text>
          </div>
          <WorkspaceBrowser onPick={start} />
        </section>
      </div>
    </div>
  );
}
