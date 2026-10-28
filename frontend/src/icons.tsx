import type { CSSProperties } from "react";

import { React, antdIcons } from "./host";

const {
  CloudUploadOutlined,
  FileExcelOutlined,
  FilePptOutlined,
  FileTextOutlined,
  FileWordOutlined,
  FolderOpenOutlined,
  FolderOutlined,
} = antdIcons;

function extension(name: string): string {
  const clean = (name || "").replace(/\\/g, "/").split("/").pop() || "";
  const dot = clean.lastIndexOf(".");
  return dot >= 0 ? clean.slice(dot + 1).toLowerCase() : "";
}

const WORD = new Set(["doc", "docm", "docx", "dot", "dotm", "dotx", "wps"]);
const SLIDES = new Set(["dps", "potx", "pps", "ppsx", "ppt", "pptm", "pptx"]);
const SHEETS = new Set([
  "csv",
  "et",
  "tsv",
  "xls",
  "xlsm",
  "xlsx",
  "xltm",
  "xltx",
]);

export function FileTypeIcon({
  name,
  directory = false,
  style,
}: {
  name: string;
  directory?: boolean;
  style?: CSSProperties;
}) {
  if (directory) {
    return <FolderOutlined style={style} />;
  }
  const ext = extension(name);
  if (WORD.has(ext)) return <FileWordOutlined style={style} />;
  if (SLIDES.has(ext)) return <FilePptOutlined style={style} />;
  if (SHEETS.has(ext)) return <FileExcelOutlined style={style} />;
  return <FileTextOutlined style={style} />;
}

export function UploadIcon() {
  return <CloudUploadOutlined />;
}

export function HeroIcon() {
  return <FileTextOutlined />;
}

export function WorkspaceIcon() {
  return <FolderOpenOutlined />;
}
