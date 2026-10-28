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
      html: "<!doctype html><html><body>DRAWER_MARKER</body></html>",
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

const { showOfficePreviewDrawer, closeOfficePreviewDrawer } = await import(
  "../src/drawer"
);

function openDrawer() {
  showOfficePreviewDrawer({
    filename: "report.xlsx",
    source: { kind: "workspace", path: "report.xlsx" },
  });
}

describe("office preview drawer", () => {
  it("slides in, renders the iframe and closes on Escape", async () => {
    openDrawer();

    await waitFor(() => {
      const drawer = document.querySelector(".qopx-drawer");
      expect(drawer).toBeTruthy();
      const frame = drawer?.querySelector(
        ".qopx-viewer-frame",
      ) as HTMLIFrameElement | null;
      expect(frame?.getAttribute("srcdoc")).toContain("DRAWER_MARKER");
    });
    expect(document.body.style.overflow).toBe("hidden");
    const drawer = document.querySelector(".qopx-drawer") as HTMLElement;
    expect(drawer.getAttribute("role")).toBe("dialog");
    expect(drawer.style.width).toMatch(/px$/);

    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));

    await waitFor(() => {
      expect(document.querySelector(".qopx-drawer")).toBeNull();
    });
    expect(document.body.style.overflow).toBe("");
  });

  it("closes when the scrim is clicked", async () => {
    openDrawer();
    await waitFor(() => {
      expect(document.querySelector(".qopx-drawer-scrim")).toBeTruthy();
    });
    const scrim = document.querySelector(".qopx-drawer-scrim") as HTMLElement;
    scrim.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    await waitFor(() => {
      expect(document.querySelector(".qopx-drawer")).toBeNull();
    });
  });

  it("can be closed programmatically", async () => {
    openDrawer();
    await waitFor(() => {
      expect(document.querySelector(".qopx-drawer")).toBeTruthy();
    });
    closeOfficePreviewDrawer();
    await waitFor(() => {
      expect(document.querySelector(".qopx-drawer")).toBeNull();
    });
  });
});
