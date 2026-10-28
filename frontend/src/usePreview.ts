import { React } from "./host";
import {
  ApiError,
  convertBlob,
  resolveSource,
  type ConvertResult,
  type PreviewRequest,
} from "./api";
import { t } from "./i18n";

export type PreviewStatus = "idle" | "loading" | "ready" | "error";

export interface PreviewState {
  status: PreviewStatus;
  filename: string;
  result?: ConvertResult;
  error?: string;
  errorCode?: string;
  blob?: Blob;
}

const { useCallback, useEffect, useRef, useState } = React;

const INITIAL: PreviewState = { status: "idle", filename: "" };

const CONVERT_TIMEOUT_MS = 240_000;

export function usePreview() {
  const [state, setState] = useState<PreviewState>(INITIAL);
  const runId = useRef(0);
  const controller = useRef<AbortController | null>(null);

  const cancel = useCallback(() => {
    runId.current += 1;
    controller.current?.abort();
    controller.current = null;
  }, []);

  const run = useCallback(async (request: PreviewRequest) => {
    const id = runId.current + 1;
    runId.current = id;
    controller.current?.abort();
    const next = new AbortController();
    controller.current = next;

    setState({ status: "loading", filename: request.filename });

    const timer = window.setTimeout(() => next.abort(), CONVERT_TIMEOUT_MS);
    try {
      const resolved = await resolveSource(request);
      if (id !== runId.current) return;
      const filename = resolved.filename || request.filename;
      setState({ status: "loading", filename });
      const result = await convertBlob(resolved.blob, filename, next.signal);
      if (id !== runId.current) return;
      setState({
        status: "ready",
        filename,
        result,
        blob: resolved.blob,
      });
    } catch (error) {
      if (id !== runId.current) return;
      if ((error as Error)?.name === "AbortError") {
        setState({
          status: "error",
          filename: request.filename,
          error: t("timeout"),
          errorCode: "aborted",
        });
        return;
      }
      const apiError = error instanceof ApiError ? error : null;
      const message =
        apiError?.status === 413
          ? t("fileTooLarge")
          : (error as Error)?.message || t("previewFailed");
      setState({
        status: "error",
        filename: request.filename,
        error: message,
        errorCode: apiError?.code,
      });
    } finally {
      window.clearTimeout(timer);
      if (controller.current === next) controller.current = null;
    }
  }, []);

  const reset = useCallback(() => {
    cancel();
    setState(INITIAL);
  }, [cancel]);

  useEffect(() => () => cancel(), [cancel]);

  return { state, run, reset };
}
