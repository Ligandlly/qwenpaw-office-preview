import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../src/drawer", () => ({
  showOfficePreviewDrawer: vi.fn(),
  closeOfficePreviewDrawer: vi.fn(),
}));

const { installPreviewInterceptor } = await import("../src/intercept");
const { showOfficePreviewDrawer } = await import("../src/drawer");

function dispatch(detail: unknown): CustomEvent {
  const event = new CustomEvent("qwenpaw:open-file-preview", {
    cancelable: true,
    detail,
  });
  window.dispatchEvent(event);
  return event;
}

describe("chat file preview interception", () => {
  beforeEach(() => {
    installPreviewInterceptor();
    vi.mocked(showOfficePreviewDrawer).mockClear();
  });

  it("intercepts office workspace targets", () => {
    const event = dispatch({
      target: {
        source: "workspace",
        path: "docs/季度报告.docx",
        root: "project",
      },
    });

    expect(event.defaultPrevented).toBe(true);
    expect(showOfficePreviewDrawer).toHaveBeenCalledWith({
      filename: "季度报告.docx",
      source: {
        kind: "workspace",
        path: "docs/季度报告.docx",
        root: "project",
      },
    });
  });

  it("intercepts artifact URLs from chat cards", () => {
    dispatch({
      target: {
        source: "attachment",
        path: "/home/u/report.pptx",
        artifactUrl: "/api/files/preview/home/u/report.pptx?token=t",
      },
    });

    expect(showOfficePreviewDrawer).toHaveBeenCalledWith({
      filename: "report.pptx",
      source: {
        kind: "artifact",
        url: "/api/files/preview/home/u/report.pptx?token=t",
        path: "/home/u/report.pptx",
      },
    });
  });

  it("stops the host listener even when it was registered first", () => {
    const hostListener = vi.fn();
    window.addEventListener("qwenpaw:open-file-preview", hostListener);
    try {
      const event = dispatch({
        target: { source: "workspace", path: "docs/a.pptx" },
      });
      expect(event.defaultPrevented).toBe(true);
      expect(hostListener).not.toHaveBeenCalled();
    } finally {
      window.removeEventListener("qwenpaw:open-file-preview", hostListener);
    }
  });

  it("treats relative attachment paths as workspace files", () => {
    dispatch({
      target: { source: "attachment", path: "docs/report.xlsx" },
    });

    expect(showOfficePreviewDrawer).toHaveBeenCalledWith({
      filename: "report.xlsx",
      source: { kind: "workspace", path: "docs/report.xlsx", root: "project" },
    });
  });

  it("ignores non-office files so the host preview keeps working", () => {
    const event = dispatch({
      target: { source: "workspace", path: "assets/logo.png" },
    });

    expect(event.defaultPrevented).toBe(false);
    expect(showOfficePreviewDrawer).not.toHaveBeenCalled();
  });

  it("ignores events without a target", () => {
    const event = dispatch({});
    expect(event.defaultPrevented).toBe(false);
    expect(showOfficePreviewDrawer).not.toHaveBeenCalled();
  });
});
