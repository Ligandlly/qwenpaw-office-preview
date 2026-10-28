import React from "react";
import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";

const previewHtml =
  '<!doctype html><html><body><h1>PREVIEW_MARKER</h1></body></html>';

vi.mock("../src/api", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../src/api")>();
  return {
    ...actual,
    fetchHealth: vi.fn(async () => ({
      ok: true,
      supported_extensions: [".docx"],
      renderers: {
        docx: true,
        pptx: true,
        xlsx: true,
        csv: true,
        legacy_doc_ppt_xls: true,
      },
      libreoffice: "/usr/bin/soffice",
      max_mb: 40,
      notes: [],
    })),
    listWorkspace: vi.fn(async () => ({
      directory: "",
      entries: [
        {
          name: "报告.docx",
          path: "报告.docx",
          kind: "file",
          size: 2048,
          modified_at: "2026-03-01T10:00:00+00:00",
          preview_kind: "binary",
        },
        {
          name: "notes.txt",
          path: "notes.txt",
          kind: "file",
          size: 20,
          modified_at: "2026-03-01T10:00:00+00:00",
          preview_kind: "text",
        },
      ],
      has_more: false,
      next_cursor: null,
    })),
    resolveSource: vi.fn(async (request: { filename: string }) => ({
      blob: new Blob(["docx-bytes"]),
      filename: request.filename,
    })),
    convertBlob: vi.fn(async () => ({
      ok: true,
      filename: "报告.docx",
      kind: "docx",
      format: "html",
      html: previewHtml,
      warnings: [],
      meta: {
        kind: "docx",
        label: "Word",
        extension: "docx",
        converted_from: null,
        size: 2048,
        html_size: previewHtml.length,
      },
    })),
  };
});

const api = await import("../src/api");
const { OfficePreviewPage } = await import("../src/OfficePreviewPage");

describe("OfficePreviewPage", () => {
  it("lists workspace documents and renders a converted preview", async () => {
    const { container } = render(<OfficePreviewPage />);

    expect(screen.getByRole("heading", { level: 1 }).textContent).toContain(
      "Office 文档预览",
    );

    const fileLabel = await screen.findByText("报告.docx");
    const row = fileLabel.closest("button");
    expect(row).toBeTruthy();
    fireEvent.click(row as HTMLButtonElement);

    await waitFor(() => {
      expect(api.convertBlob).toHaveBeenCalledTimes(1);
    });
    const call = vi.mocked(api.convertBlob).mock.calls[0];
    expect(call[1]).toBe("报告.docx");

    await waitFor(() => {
      const frame = container.querySelector(
        ".qopx-viewer-frame",
      ) as HTMLIFrameElement | null;
      expect(frame).toBeTruthy();
      expect(frame?.getAttribute("srcdoc")).toContain("PREVIEW_MARKER");
      expect(frame?.getAttribute("sandbox")).toBe("");
    });
  });

  it("hides unsupported files when the filter is on", async () => {
    render(<OfficePreviewPage />);
    await screen.findByText("报告.docx");
    expect(screen.queryByText("notes.txt")).toBeNull();
  });

  it("accepts dropped office files", async () => {
    vi.mocked(api.convertBlob).mockClear();
    const { container } = render(<OfficePreviewPage />);
    const zone = container.querySelector(".qopx-drop") as HTMLElement;
    const file = new File(["x"], "deck.pptx", {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    fireEvent.drop(zone, { dataTransfer: { files: [file] } });

    await waitFor(() => {
      expect(api.convertBlob).toHaveBeenCalledTimes(1);
    });
    expect(vi.mocked(api.convertBlob).mock.calls[0][1]).toBe("deck.pptx");
  });
});
