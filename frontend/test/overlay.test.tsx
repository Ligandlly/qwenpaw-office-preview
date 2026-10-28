import { waitFor } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";

vi.mock("../src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../src/api")>();
  return {
    ...actual,
    resolveSource: vi.fn(async () => ({
      blob: new Blob(["bytes"]),
      filename: "report.xlsx",
    })),
    convertBlob: vi.fn(async () => ({
      ok: true,
      filename: "report.xlsx",
      kind: "xlsx",
      format: "html",
      html: "<!doctype html><html><body>OVERLAY_MARKER</body></html>",
      warnings: ["row truncated"],
      meta: {
        kind: "xlsx",
        label: "Excel",
        extension: "xlsx",
        converted_from: null,
        size: 10,
        html_size: 60,
      },
    })),
  };
});

const { showOfficePreviewOverlay, closeOfficePreviewOverlay } = await import(
  "../src/overlay"
);

describe("office preview overlay", () => {
  it("opens, renders the iframe and closes on Escape", async () => {
    showOfficePreviewOverlay({
      filename: "report.xlsx",
      source: { kind: "workspace", path: "report.xlsx" },
    });

    await waitFor(() => {
      const overlay = document.querySelector(".qopx-overlay");
      expect(overlay).toBeTruthy();
      const frame = overlay?.querySelector(
        ".qopx-viewer-frame",
      ) as HTMLIFrameElement | null;
      expect(frame?.getAttribute("srcdoc")).toContain("OVERLAY_MARKER");
    });
    expect(document.body.style.overflow).toBe("hidden");

    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));

    await waitFor(() => {
      expect(document.querySelector(".qopx-overlay")).toBeNull();
    });
    expect(document.body.style.overflow).toBe("");
  });

  it("can be closed programmatically", async () => {
    showOfficePreviewOverlay({
      filename: "report.xlsx",
      source: { kind: "workspace", path: "report.xlsx" },
    });
    await waitFor(() => {
      expect(document.querySelector(".qopx-overlay")).toBeTruthy();
    });
    closeOfficePreviewOverlay();
    await waitFor(() => {
      expect(document.querySelector(".qopx-overlay")).toBeNull();
    });
  });
});
