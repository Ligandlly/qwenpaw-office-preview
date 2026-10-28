import { beforeEach, describe, expect, it, vi } from "vitest";

import {
  ApiError,
  convertBlob,
  extensionOf,
  isOfficeFile,
  resolveSource,
} from "../src/api";

const host = window.QwenPaw.host;

function jsonResponse(payload: unknown, status = 200): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

describe("extension detection", () => {
  it("recognises office extensions case-insensitively", () => {
    expect(extensionOf("a/b/Report.DOCX")).toBe("docx");
    expect(isOfficeFile("/tmp/演示.pptx?token=1")).toBe(true);
    expect(isOfficeFile("notes.txt")).toBe(false);
    expect(extensionOf("no-extension")).toBe("");
  });
});

describe("resolveSource", () => {
  beforeEach(() => {
    host.fetch = vi.fn();
  });

  it("fetches workspace files through the authenticated host API", async () => {
    const fetchMock = host.fetch as unknown as ReturnType<typeof vi.fn>;
    // A string body works across Node/undici versions; constructing a
    // Response from a jsdom Blob fails on older Node runtimes.
    fetchMock.mockResolvedValue(new Response("data", { status: 200 }));

    const resolved = await resolveSource({
      filename: "report.docx",
      source: { kind: "workspace", path: "docs/report.docx" },
    });

    expect(resolved.filename).toBe("report.docx");
    const url = fetchMock.mock.calls[0][0] as string;
    expect(url).toContain("/workspace/file-download?");
    expect(url).toContain("path=docs%2Freport.docx");
    expect(url).toContain("root=project");
  });

  it("passes uploads through unchanged", async () => {
    const blob = new Blob(["upload"]);
    const resolved = await resolveSource({
      filename: "local.xlsx",
      source: { kind: "upload", blob },
    });
    expect(resolved.blob).toBe(blob);
    expect(resolved.filename).toBe("local.xlsx");
  });
});

describe("convertBlob", () => {
  beforeEach(() => {
    host.fetch = vi.fn();
  });

  it("posts bytes and returns the parsed result", async () => {
    const fetchMock = host.fetch as unknown as ReturnType<typeof vi.fn>;
    fetchMock.mockResolvedValue(
      jsonResponse({
        ok: true,
        filename: "a.docx",
        kind: "docx",
        format: "html",
        html: "<html>ok</html>",
        warnings: [],
        meta: {
          kind: "docx",
          label: "Word",
          extension: "docx",
          converted_from: null,
          size: 3,
          html_size: 10,
        },
      }),
    );

    const result = await convertBlob(new Blob(["abc"]), "a.docx");
    expect(result.html).toContain("ok");
    const url = fetchMock.mock.calls[0][0] as string;
    expect(url).toContain("/office-preview/convert?");
    expect(url).toContain("filename=a.docx");
    expect(url).toContain("theme=");
  });

  it("passes the current host theme to the converter", async () => {
    const fetchMock = host.fetch as unknown as ReturnType<typeof vi.fn>;
    fetchMock.mockResolvedValue(
      jsonResponse({
        ok: true,
        filename: "a.docx",
        kind: "docx",
        format: "html",
        html: "<html>ok</html>",
        warnings: [],
        meta: {},
      }),
    );
    document.documentElement.classList.add("dark-mode");
    try {
      await convertBlob(new Blob(["abc"]), "a.docx");
      const url = fetchMock.mock.calls[0][0] as string;
      expect(url).toContain("theme=dark");
    } finally {
      document.documentElement.classList.remove("dark-mode");
    }
  });

  it("surfaces structured backend errors", async () => {
    const fetchMock = host.fetch as unknown as ReturnType<typeof vi.fn>;
    fetchMock.mockResolvedValue(
      jsonResponse(
        {
          detail: {
            code: "unsupported_format",
            message: "不支持预览该文件类型",
          },
        },
        415,
      ),
    );

    await expect(convertBlob(new Blob(["x"]), "a.png")).rejects.toMatchObject({
      name: "ApiError",
      status: 415,
      code: "unsupported_format",
    } satisfies Partial<ApiError>);
  });
});
