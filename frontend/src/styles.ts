const STYLE_ID = "qwenpaw-office-preview-styles";

const CSS = `
.qopx-page,
.qopx-drawer {
  --qopx-ink: #1f2328;
  --qopx-muted: #5b6472;
  --qopx-bg: #ffffff;
  --qopx-subtle: #fafbfc;
  --qopx-subtle-2: #f1f3f5;
  --qopx-line: #e5e7eb;
  --qopx-line-soft: #eef0f3;
  --qopx-accent: #2563eb;
  --qopx-accent-soft: rgba(37, 99, 235, .12);
  --qopx-frame: #eef1f5;
  --qopx-dots: rgba(148, 163, 184, .18);
  --qopx-shadow: 0 1px 2px rgba(15, 23, 42, .04);
  --qopx-drawer-shadow: -24px 0 56px rgba(15, 23, 42, .18),
    -2px 0 10px rgba(15, 23, 42, .06);
  color-scheme: light;
}
.qopx-page[data-theme="dark"],
.qopx-drawer[data-theme="dark"] {
  --qopx-ink: #e6e9ef;
  --qopx-muted: #a0a8b6;
  --qopx-bg: #1c2027;
  --qopx-subtle: #22272f;
  --qopx-subtle-2: #2a3039;
  --qopx-line: #313743;
  --qopx-line-soft: #2a3039;
  --qopx-accent: #7aa2ff;
  --qopx-accent-soft: rgba(122, 162, 255, .16);
  --qopx-frame: #14171c;
  --qopx-dots: rgba(148, 163, 184, .12);
  --qopx-shadow: 0 1px 2px rgba(0, 0, 0, .35);
  --qopx-drawer-shadow: -24px 0 56px rgba(0, 0, 0, .55),
    -2px 0 10px rgba(0, 0, 0, .30);
  color-scheme: dark;
}

.qopx-page {
  max-width: 1180px;
  margin: 0 auto;
  padding: 22px 24px 40px;
  color: var(--qopx-ink);
  font-size: 14px;
}
.qopx-hero { margin-bottom: 18px; }
.qopx-hero h1 {
  display: flex; align-items: center; gap: 10px;
  margin: 0 0 6px; font-size: 22px; font-weight: 650;
  color: var(--qopx-ink);
}
.qopx-hero p { margin: 0; color: var(--qopx-muted); line-height: 1.7; max-width: 760px; }
.qopx-hero-icon {
  display: inline-flex; align-items: center; justify-content: center;
  width: 36px; height: 36px; border-radius: 10px;
  background: var(--qopx-accent-soft);
  color: var(--qopx-accent); font-size: 18px;
}
.qopx-grid { display: grid; grid-template-columns: minmax(0, 1fr); gap: 18px; }
.qopx-card {
  background: var(--qopx-bg);
  border: 1px solid var(--qopx-line);
  border-radius: 12px;
  box-shadow: var(--qopx-shadow);
  overflow: hidden;
  color: var(--qopx-ink);
}
.qopx-card-head {
  display: flex; align-items: center; justify-content: space-between;
  gap: 12px; padding: 12px 16px;
  border-bottom: 1px solid var(--qopx-line-soft);
}
.qopx-card-title { font-weight: 600; display: flex; align-items: center; gap: 8px; }
.qopx-card-body { padding: 16px; }

.qopx-drop {
  display: flex; flex-direction: column; align-items: center; justify-content: center;
  gap: 8px; width: 100%; min-height: 168px; padding: 24px;
  border: 1.5px dashed var(--qopx-line);
  border-radius: 12px; cursor: pointer; text-align: center;
  color: var(--qopx-ink);
  background:
    radial-gradient(circle at 1px 1px, var(--qopx-dots) 1px, transparent 0)
    0 0 / 16px 16px,
    var(--qopx-subtle);
  transition: border-color .18s ease, background-color .18s ease, transform .18s ease;
}
.qopx-drop:hover { border-color: var(--qopx-accent); }
.qopx-drop.qopx-dragging {
  border-color: var(--qopx-accent);
  background-color: var(--qopx-accent-soft);
  transform: translateY(-1px);
}
.qopx-drop-icon {
  display: inline-flex; align-items: center; justify-content: center;
  font-size: 34px; line-height: 1; color: var(--qopx-accent);
}
.qopx-drop-title { font-size: 15px; font-weight: 600; }
.qopx-drop-hint { font-size: 12.5px; color: var(--qopx-muted); }
.qopx-drop-actions { display: flex; gap: 10px; margin-top: 6px; }

.qopx-status { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 14px; }
.qopx-chip {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 3px 10px; border-radius: 999px; font-size: 12px;
  border: 1px solid var(--qopx-line);
  background: var(--qopx-subtle);
  color: var(--qopx-ink);
}
.qopx-dot { width: 7px; height: 7px; border-radius: 50%; background: #22c55e; }
.qopx-dot.qopx-bad { background: #ef4444; }
.qopx-dot.qopx-muted { background: #94a3b8; }

.qopx-toolbar { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.qopx-breadcrumb {
  display: flex; align-items: center; gap: 4px; flex-wrap: wrap;
  padding: 10px 16px; border-bottom: 1px solid var(--qopx-line-soft);
  font-size: 13px;
}
.qopx-crumb {
  border: 0; background: transparent; padding: 2px 6px; border-radius: 6px;
  cursor: pointer; color: inherit; opacity: .82;
}
.qopx-crumb:hover { background: var(--qopx-subtle-2); opacity: 1; }
.qopx-crumb-sep { opacity: .35; }
.qopx-file-list { max-height: 46vh; overflow: auto; }
.qopx-file-row {
  display: flex; align-items: center; gap: 10px; width: 100%;
  padding: 9px 16px; border: 0; border-bottom: 1px solid var(--qopx-line-soft);
  background: transparent; cursor: pointer; text-align: left; color: inherit;
  font-size: 13.5px;
}
.qopx-file-row:hover:not(:disabled) { background: var(--qopx-subtle); }
.qopx-file-row:disabled { cursor: not-allowed; opacity: .45; }
.qopx-file-icon {
  width: 30px; height: 30px; flex: 0 0 30px; border-radius: 8px;
  display: inline-flex; align-items: center; justify-content: center;
  font-size: 16px; background: var(--qopx-subtle-2);
  color: var(--qopx-muted);
}
.qopx-file-main { min-width: 0; flex: 1; }
.qopx-file-name { display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.qopx-file-sub { display: block; font-size: 11.5px; color: var(--qopx-muted); }
.qopx-file-tag { font-size: 11px; color: var(--qopx-muted); white-space: nowrap; }
.qopx-empty { padding: 26px 16px; text-align: center; color: var(--qopx-muted); }

.qopx-viewer { display: flex; flex-direction: column; min-height: 0; color: var(--qopx-ink); }
.qopx-viewer-head {
  display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
  padding: 10px 14px; border-bottom: 1px solid var(--qopx-line-soft);
  color: var(--qopx-ink);
}
.qopx-viewer-title {
  display: flex; align-items: center; gap: 8px; min-width: 0; flex: 1;
  font-weight: 600;
}
.qopx-viewer-title span {
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.qopx-viewer-file-icon {
  display: inline-flex; align-items: center; color: var(--qopx-muted);
  font-size: 16px;
}
.qopx-viewer-actions { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
.qopx-viewer-frame {
  flex: 1; min-height: 0; width: 100%; border: 0;
  background: var(--qopx-frame); display: block;
}
.qopx-viewer-state {
  display: flex; flex-direction: column; align-items: center; justify-content: center;
  gap: 10px; padding: 42px 20px; text-align: center; color: var(--qopx-muted);
}

.qopx-drawer-scrim {
  position: fixed; inset: 0; z-index: 2147483000;
  background: transparent;
}
.qopx-drawer-scrim.qopx-closing { pointer-events: none; }
.qopx-drawer {
  position: fixed; top: 0; right: 0; bottom: 0;
  z-index: 2147483001;
  display: flex; flex-direction: column;
  min-width: 440px;
  background: var(--qopx-bg);
  color: var(--qopx-ink);
  border-left: 1px solid var(--qopx-line);
  box-shadow: var(--qopx-drawer-shadow);
  animation: qopx-drawer-in .22s cubic-bezier(.22, .78, .24, 1);
  will-change: transform;
  outline: none;
}
.qopx-drawer.qopx-closing {
  transform: translateX(102%);
  transition: transform .18s ease-in;
}
@keyframes qopx-drawer-in {
  from { transform: translateX(102%); }
  to { transform: translateX(0); }
}
.qopx-drawer .qopx-viewer { flex: 1; min-height: 0; }
.qopx-drawer .qopx-viewer-head {
  min-height: 58px; padding: 11px 16px;
  background: var(--qopx-bg);
  background: color-mix(in srgb, var(--qopx-bg) 92%, transparent);
  backdrop-filter: blur(14px);
}
.qopx-drawer-resizer {
  position: absolute; left: -4px; top: 0; bottom: 0; width: 9px;
  cursor: col-resize; z-index: 3; outline: none;
}
.qopx-drawer-resizer::after {
  content: ""; position: absolute; left: 3px; top: 50%;
  width: 3px; height: 42px; transform: translateY(-50%);
  border-radius: 3px; background: var(--qopx-line);
  opacity: 0; transition: opacity .16s ease;
}
.qopx-drawer:hover .qopx-drawer-resizer::after,
.qopx-drawer-resizer:focus-visible::after { opacity: .65; }
.qopx-drawer-scrim.qopx-closing + .qopx-drawer { pointer-events: none; }
.qopx-viewer-body { flex: 1; min-height: 0; display: flex; flex-direction: column; }
.qopx-warning-list { margin: 0; padding-left: 18px; }

@media (max-width: 720px) {
  .qopx-page { padding: 14px 12px 28px; }
  .qopx-drawer {
    width: 100vw !important; min-width: 0;
    border-left: 0;
  }
  .qopx-drawer-resizer { display: none; }
  .qopx-drawer-scrim { background: rgba(0, 0, 0, .38); }
  .qopx-file-list { max-height: 52vh; }
}
`;

export function installStyles(): void {
  if (typeof document === "undefined") return;
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.appendChild(style);
}
