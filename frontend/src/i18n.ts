const ZH = {
  routeLabel: "Office 预览",
  routeTitle: "Office 文档预览",
  pageSubtitle:
    "无需下载即可查看 Word、PowerPoint 和 Excel 文档。拖入本地文件，或直接从当前工作区打开。",
  dropTitle: "拖拽文件到这里，或点击选择",
  dropHint: "支持 doc / docx / ppt / pptx / xls / xlsx，以及 csv / tsv",
  chooseFile: "选择文件",
  workspace: "工作区文件",
  projectRoot: "项目目录",
  workspaceRoot: "工作区目录",
  refresh: "刷新",
  filterOnly: "仅显示可预览",
  emptyDir: "这个目录里没有可预览的文档。",
  loadDirFailed: "读取目录失败",
  unsupported: "不支持预览该文件类型",
  unsupportedHint: "支持 Word、PowerPoint、Excel 文档以及 CSV。",
  loading: "正在生成预览…",
  converting: "正在解析文档，较大的文件可能需要几秒钟。",
  retry: "重试",
  close: "关闭",
  download: "下载原文件",
  openWindow: "新窗口打开",
  previewFailed: "预览失败",
  backToList: "返回文件列表",
  size: "大小",
  modified: "修改时间",
  status: "转换器状态",
  statusOk: "可用",
  statusMissing: "未安装",
  healthFailed: "无法获取转换器状态",
  depsMissing: "部分转换器依赖缺失，请重新安装插件以自动补齐依赖。",
  legacyHint:
    "旧版 .doc / .ppt / .xls 文件需要本机安装 LibreOffice；也可以先另存为 docx / pptx / xlsx。",
  cacheNote: "转换结果会缓存在内存中，重复打开同一文件会更快。",
  fileTooLarge: "文件过大，无法在线预览。",
  dialogLabel: "Office 文档预览",
  navigate: "定位",
  parent: "上一级",
  root: "根目录",
  entries: "项",
  openInPage: "在独立页面中打开",
  emptyState: "选择左侧文件或拖入文档开始预览",
  convertedFrom: "已从旧格式转换",
  warnings: "格式提示",
  timeout: "预览已取消或超时。",
  resizeHint: "拖动调整宽度，双击恢复默认",
};

const EN: Record<keyof typeof ZH, string> = {
  routeLabel: "Office Preview",
  routeTitle: "Office Document Preview",
  pageSubtitle:
    "Read Word, PowerPoint and Excel documents without downloading them. Drop a local file or open one from the workspace.",
  dropTitle: "Drop a file here, or click to choose",
  dropHint: "Supports doc / docx / ppt / pptx / xls / xlsx, plus csv / tsv",
  chooseFile: "Choose file",
  workspace: "Workspace files",
  projectRoot: "Project",
  workspaceRoot: "Workspace",
  refresh: "Refresh",
  filterOnly: "Previewable only",
  emptyDir: "No previewable documents in this folder.",
  loadDirFailed: "Failed to read the folder",
  unsupported: "This file type cannot be previewed",
  unsupportedHint: "Word, PowerPoint, Excel and CSV documents are supported.",
  loading: "Building preview…",
  converting: "Parsing the document; large files may take a few seconds.",
  retry: "Retry",
  close: "Close",
  download: "Download original",
  openWindow: "Open in new window",
  previewFailed: "Preview failed",
  backToList: "Back to files",
  size: "Size",
  modified: "Modified",
  status: "Converter status",
  statusOk: "Available",
  statusMissing: "Not installed",
  healthFailed: "Could not read converter status",
  depsMissing:
    "Some converters are unavailable. Reinstall the plugin to restore its Python dependencies.",
  legacyHint:
    "Legacy .doc / .ppt / .xls previews require LibreOffice on the host; alternatively save the file as docx / pptx / xlsx.",
  cacheNote:
    "Converted documents are cached in memory, so reopening the same file is faster.",
  fileTooLarge: "The file is too large to preview online.",
  dialogLabel: "Office document preview",
  navigate: "Go",
  parent: "Up",
  root: "Root",
  entries: "items",
  openInPage: "Open in dedicated page",
  emptyState: "Pick a file on the left or drop a document to begin",
  convertedFrom: "Converted from a legacy format",
  warnings: "Document notes",
  timeout: "Preview was cancelled or timed out.",
  resizeHint: "Drag to resize, double-click to reset",
};

export type MessageKey = keyof typeof ZH;

export function currentLocale(): "zh" | "en" {
  try {
    const stored =
      window.localStorage.getItem("language") ||
      document.documentElement.getAttribute("lang") ||
      navigator.language ||
      "en";
    return stored.toLowerCase().startsWith("zh") ? "zh" : "en";
  } catch {
    return "zh";
  }
}

export function t(key: MessageKey): string {
  return (currentLocale() === "zh" ? ZH : EN)[key];
}

export function localizedLabel(
  zh: string,
  en: string,
): (locale: string) => string {
  return (locale: string) =>
    String(locale || "").toLowerCase().startsWith("zh") ? zh : en;
}
