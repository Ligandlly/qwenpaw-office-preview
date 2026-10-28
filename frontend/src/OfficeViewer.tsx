import type { ReactNode } from "react";

import { React, antd, antdIcons } from "./host";
import { t } from "./i18n";
import { formatBytes } from "./format";
import { FileTypeIcon } from "./icons";
import type { PreviewState } from "./usePreview";

const { Alert, Button, Space, Spin, Tag, Tooltip, Typography } = antd;
const {
  CloseOutlined,
  DownloadOutlined,
  ExportOutlined,
  FileWarningOutlined,
  ReloadOutlined,
} = antdIcons;

export interface OfficeViewerProps {
  state: PreviewState;
  onClose?: () => void;
  onRetry?: () => void;
  height?: number | string;
  className?: string;
}

function triggerDownload(blob: Blob | undefined, filename: string): void {
  if (!blob) return;
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename || "document";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

function openInNewWindow(html: string): void {
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const opened = window.open(url, "_blank", "noopener,noreferrer");
  if (!opened) {
    URL.revokeObjectURL(url);
    return;
  }
  window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

export function OfficeViewer({
  state,
  onClose,
  onRetry,
  height = "72vh",
  className,
}: OfficeViewerProps) {
  const { status, result, filename, error } = state;
  const kind = result?.meta?.label || result?.kind?.toUpperCase() || "";
  const warnings = result?.warnings ?? [];
  const convertedFrom = result?.meta?.converted_from;

  const header = (
    <div className="qopx-viewer-head">
      <div className="qopx-viewer-title">
        <span className="qopx-viewer-file-icon" aria-hidden="true">
          <FileTypeIcon name={filename} />
        </span>
        <span title={filename}>{filename || t("routeTitle")}</span>
        {status === "ready" && kind && <Tag color="blue">{kind}</Tag>}
        {convertedFrom && (
          <Tooltip title={t("convertedFrom")}>
            <Tag>{String(convertedFrom).replace(".", "").toUpperCase()}</Tag>
          </Tooltip>
        )}
        {result?.cached && <Tag color="default">cache</Tag>}
        {result?.meta?.size ? (
          <Typography.Text type="secondary" style={{ fontSize: 12 }}>
            {formatBytes(result.meta.size)}
          </Typography.Text>
        ) : null}
      </div>
      <Space size={6} wrap>
        {status === "ready" && result && (
          <>
            {state.blob && (
              <Tooltip title={t("download")}>
                <Button
                  size="small"
                  aria-label={t("download")}
                  icon={<DownloadOutlined />}
                  onClick={() => triggerDownload(state.blob as Blob, filename)}
                />
              </Tooltip>
            )}
            <Tooltip title={t("openWindow")}>
              <Button
                size="small"
                aria-label={t("openWindow")}
                icon={<ExportOutlined />}
                onClick={() => openInNewWindow(result.html)}
              />
            </Tooltip>
          </>
        )}
        {onRetry && status !== "loading" && (
          <Tooltip title={t("retry")}>
            <Button
              size="small"
              aria-label={t("retry")}
              icon={<ReloadOutlined />}
              onClick={onRetry}
            />
          </Tooltip>
        )}
        {onClose && (
          <Tooltip title={t("close")}>
            <Button
              size="small"
              type="text"
              aria-label={t("close")}
              icon={<CloseOutlined />}
              onClick={onClose}
            />
          </Tooltip>
        )}
      </Space>
    </div>
  );

  let body: ReactNode;
  if (status === "loading") {
    body = (
      <div className="qopx-viewer-state">
        <Spin size="large" />
        <Typography.Text strong>{t("loading")}</Typography.Text>
        <Typography.Text type="secondary">{t("converting")}</Typography.Text>
      </div>
    );
  } else if (status === "error") {
    body = (
      <div style={{ padding: 22 }}>
        <Alert
          type="error"
          showIcon
          icon={<FileWarningOutlined />}
          message={t("previewFailed")}
          description={
            <div>
              <p style={{ margin: "0 0 10px" }}>{error}</p>
              {state.errorCode === "legacy_converter_missing" && (
                <p style={{ margin: 0, opacity: 0.75 }}>{t("legacyHint")}</p>
              )}
              {state.errorCode === "missing_dependency" && (
                <p style={{ margin: 0, opacity: 0.75 }}>{t("depsMissing")}</p>
              )}
              {onRetry && (
                <Button
                  size="small"
                  style={{ marginTop: 12 }}
                  icon={<ReloadOutlined />}
                  onClick={onRetry}
                >
                  {t("retry")}
                </Button>
              )}
            </div>
          }
        />
      </div>
    );
  } else if (status === "ready" && result) {
    body = (
      <>
        {warnings.length > 0 && (
          <Alert
            type="warning"
            showIcon
            style={{ borderRadius: 0, borderInline: 0, borderTop: 0 }}
            message={t("warnings")}
            description={
              <ul className="qopx-warning-list">
                {warnings.slice(0, 6).map((warning, index) => (
                  <li key={`${index}-${warning.slice(0, 24)}`}>{warning}</li>
                ))}
              </ul>
            }
          />
        )}
        <iframe
          className="qopx-viewer-frame"
          title={filename || t("routeTitle")}
          srcDoc={result.html}
          sandbox=""
          referrerPolicy="no-referrer"
        />
      </>
    );
  } else {
    body = (
      <div className="qopx-viewer-state">
        <Typography.Text type="secondary">{t("emptyState")}</Typography.Text>
      </div>
    );
  }

  return (
    <div
      className={`qopx-viewer${className ? ` ${className}` : ""}`}
      style={{ height }}
    >
      {header}
      <div className="qopx-viewer-body">{body}</div>
    </div>
  );
}
