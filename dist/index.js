var w=window.QwenPaw.host,e=w.React,ee=w.ReactDOM,b=w.antd,I=w.antdIcons;function k(){try{if(document.documentElement.classList.contains("dark-mode"))return"dark";w.useTheme}catch{}return"light"}function Pe(){try{let t=window.sessionStorage.getItem("qwenpaw-agent-storage")??window.localStorage.getItem("qwenpaw-agent-storage");if(t){let o=JSON.parse(t)?.state?.selectedAgent;if(typeof o=="string"&&o)return o}}catch{}try{return w.getSelectedAgentId?.()??"default"}catch{return"default"}}function te(){let t={};try{let o=w.getApiToken();o&&(t.Authorization=`Bearer ${o}`)}catch{}let r=Pe();return r&&(t["X-Agent-Id"]=r),t}async function S(t,r={}){if(w.fetch)return w.fetch(t,r);let o=/^https?:\/\//i.test(t)?t:w.getApiUrl(t);return fetch(o,{...r,headers:{...te(),...r.headers??{}}})}async function re(t,r={}){return fetch(t,{...r,headers:{...te(),...r.headers??{}}})}var oe=new Set(["csv","doc","docm","docx","dot","dotm","dotx","dps","et","pps","ppsx","ppt","pptm","pptx","potx","tsv","wps","xls","xlsm","xlsx","xltm","xltx"]),ne=Array.from(oe).map(t=>`.${t}`).join(","),q=class extends Error{constructor(r,o,n){super(r),this.name="ApiError",this.status=o,this.code=n}};function Te(t){let o=(t||"").split(/[?#]/,1)[0].replace(/\\/g,"/").split("/").pop()||"",n=o.lastIndexOf(".");return n>=0?o.slice(n+1).toLowerCase():""}function P(t){return oe.has(Te(t))}function T(t){return(t||"").replace(/\\/g,"/").split(/[?#]/,1)[0].split("/").filter(Boolean).pop()||"document"}async function H(t){let r=t.status,o=`${r} ${t.statusText}`.trim(),n=`http_${r}`;try{let s=await t.json(),a=s?.detail??s;typeof a=="string"?o=a:a&&typeof a=="object"&&(typeof a.message=="string"&&(o=a.message),typeof a.code=="string"&&(n=a.code))}catch{try{let s=await t.text();s&&(o=s.slice(0,500))}catch{}}return new q(o,r,n)}async function ie(){let t=await S("/office-preview/health");if(!t.ok)throw await H(t);return await t.json()}async function ae(t,r,o=400){let n=new URLSearchParams({path:t,root:r,limit:String(o)}),s=await S(`/workspace/tree?${n.toString()}`);if(!s.ok)throw await H(s);return await s.json()}async function se(t){let{source:r}=t;if(r.kind==="upload")return{blob:r.blob,filename:t.filename};let o;if(r.kind==="workspace"){let a=new URLSearchParams({path:r.path,root:r.root||"project"});o=await S(`/workspace/file-download?${a.toString()}`)}else o=await re(r.url);if(!o.ok)throw await H(o);let n=await o.blob(),s=t.filename||T(r.kind==="workspace"?r.path:r.path||"");return{blob:n,filename:s}}async function le(t,r,o){let n=new URLSearchParams({filename:r||"document.docx",theme:k()}),s=await S(`/office-preview/convert?${n.toString()}`,{method:"POST",headers:{"Content-Type":"application/octet-stream"},body:t,signal:o});if(!s.ok)throw await H(s);let a=await s.json();if(!a||typeof a.html!="string"||!a.html)throw new q("The server returned an empty preview.",500,"empty_result");return a}var Ee={routeLabel:"Office \u9884\u89C8",routeTitle:"Office \u6587\u6863\u9884\u89C8",pageSubtitle:"\u65E0\u9700\u4E0B\u8F7D\u5373\u53EF\u67E5\u770B Word\u3001PowerPoint \u548C Excel \u6587\u6863\u3002\u62D6\u5165\u672C\u5730\u6587\u4EF6\uFF0C\u6216\u76F4\u63A5\u4ECE\u5F53\u524D\u5DE5\u4F5C\u533A\u6253\u5F00\u3002",dropTitle:"\u62D6\u62FD\u6587\u4EF6\u5230\u8FD9\u91CC\uFF0C\u6216\u70B9\u51FB\u9009\u62E9",dropHint:"\u652F\u6301 doc / docx / ppt / pptx / xls / xlsx\uFF0C\u4EE5\u53CA csv / tsv",chooseFile:"\u9009\u62E9\u6587\u4EF6",workspace:"\u5DE5\u4F5C\u533A\u6587\u4EF6",projectRoot:"\u9879\u76EE\u76EE\u5F55",workspaceRoot:"\u5DE5\u4F5C\u533A\u76EE\u5F55",refresh:"\u5237\u65B0",filterOnly:"\u4EC5\u663E\u793A\u53EF\u9884\u89C8",emptyDir:"\u8FD9\u4E2A\u76EE\u5F55\u91CC\u6CA1\u6709\u53EF\u9884\u89C8\u7684\u6587\u6863\u3002",loadDirFailed:"\u8BFB\u53D6\u76EE\u5F55\u5931\u8D25",unsupported:"\u4E0D\u652F\u6301\u9884\u89C8\u8BE5\u6587\u4EF6\u7C7B\u578B",unsupportedHint:"\u652F\u6301 Word\u3001PowerPoint\u3001Excel \u6587\u6863\u4EE5\u53CA CSV\u3002",loading:"\u6B63\u5728\u751F\u6210\u9884\u89C8\u2026",converting:"\u6B63\u5728\u89E3\u6790\u6587\u6863\uFF0C\u8F83\u5927\u7684\u6587\u4EF6\u53EF\u80FD\u9700\u8981\u51E0\u79D2\u949F\u3002",retry:"\u91CD\u8BD5",close:"\u5173\u95ED",download:"\u4E0B\u8F7D\u539F\u6587\u4EF6",openWindow:"\u65B0\u7A97\u53E3\u6253\u5F00",previewFailed:"\u9884\u89C8\u5931\u8D25",backToList:"\u8FD4\u56DE\u6587\u4EF6\u5217\u8868",size:"\u5927\u5C0F",modified:"\u4FEE\u6539\u65F6\u95F4",status:"\u8F6C\u6362\u5668\u72B6\u6001",statusOk:"\u53EF\u7528",statusMissing:"\u672A\u5B89\u88C5",healthFailed:"\u65E0\u6CD5\u83B7\u53D6\u8F6C\u6362\u5668\u72B6\u6001",depsMissing:"\u90E8\u5206\u8F6C\u6362\u5668\u4F9D\u8D56\u7F3A\u5931\uFF0C\u8BF7\u91CD\u65B0\u5B89\u88C5\u63D2\u4EF6\u4EE5\u81EA\u52A8\u8865\u9F50\u4F9D\u8D56\u3002",legacyHint:"\u65E7\u7248 .doc / .ppt / .xls \u6587\u4EF6\u9700\u8981\u672C\u673A\u5B89\u88C5 LibreOffice\uFF1B\u4E5F\u53EF\u4EE5\u5148\u53E6\u5B58\u4E3A docx / pptx / xlsx\u3002",cacheNote:"\u8F6C\u6362\u7ED3\u679C\u4F1A\u7F13\u5B58\u5728\u5185\u5B58\u4E2D\uFF0C\u91CD\u590D\u6253\u5F00\u540C\u4E00\u6587\u4EF6\u4F1A\u66F4\u5FEB\u3002",fileTooLarge:"\u6587\u4EF6\u8FC7\u5927\uFF0C\u65E0\u6CD5\u5728\u7EBF\u9884\u89C8\u3002",dialogLabel:"Office \u6587\u6863\u9884\u89C8",navigate:"\u5B9A\u4F4D",parent:"\u4E0A\u4E00\u7EA7",root:"\u6839\u76EE\u5F55",entries:"\u9879",openInPage:"\u5728\u72EC\u7ACB\u9875\u9762\u4E2D\u6253\u5F00",emptyState:"\u9009\u62E9\u5DE6\u4FA7\u6587\u4EF6\u6216\u62D6\u5165\u6587\u6863\u5F00\u59CB\u9884\u89C8",convertedFrom:"\u5DF2\u4ECE\u65E7\u683C\u5F0F\u8F6C\u6362",warnings:"\u683C\u5F0F\u63D0\u793A",timeout:"\u9884\u89C8\u5DF2\u53D6\u6D88\u6216\u8D85\u65F6\u3002",resizeHint:"\u62D6\u52A8\u8C03\u6574\u5BBD\u5EA6\uFF0C\u53CC\u51FB\u6062\u590D\u9ED8\u8BA4"},Se={routeLabel:"Office Preview",routeTitle:"Office Document Preview",pageSubtitle:"Read Word, PowerPoint and Excel documents without downloading them. Drop a local file or open one from the workspace.",dropTitle:"Drop a file here, or click to choose",dropHint:"Supports doc / docx / ppt / pptx / xls / xlsx, plus csv / tsv",chooseFile:"Choose file",workspace:"Workspace files",projectRoot:"Project",workspaceRoot:"Workspace",refresh:"Refresh",filterOnly:"Previewable only",emptyDir:"No previewable documents in this folder.",loadDirFailed:"Failed to read the folder",unsupported:"This file type cannot be previewed",unsupportedHint:"Word, PowerPoint, Excel and CSV documents are supported.",loading:"Building preview\u2026",converting:"Parsing the document; large files may take a few seconds.",retry:"Retry",close:"Close",download:"Download original",openWindow:"Open in new window",previewFailed:"Preview failed",backToList:"Back to files",size:"Size",modified:"Modified",status:"Converter status",statusOk:"Available",statusMissing:"Not installed",healthFailed:"Could not read converter status",depsMissing:"Some converters are unavailable. Reinstall the plugin to restore its Python dependencies.",legacyHint:"Legacy .doc / .ppt / .xls previews require LibreOffice on the host; alternatively save the file as docx / pptx / xlsx.",cacheNote:"Converted documents are cached in memory, so reopening the same file is faster.",fileTooLarge:"The file is too large to preview online.",dialogLabel:"Office document preview",navigate:"Go",parent:"Up",root:"Root",entries:"items",openInPage:"Open in dedicated page",emptyState:"Pick a file on the left or drop a document to begin",convertedFrom:"Converted from a legacy format",warnings:"Document notes",timeout:"Preview was cancelled or timed out.",resizeHint:"Drag to resize, double-click to reset"};function De(){try{return(window.localStorage.getItem("language")||document.documentElement.getAttribute("lang")||navigator.language||"en").toLowerCase().startsWith("zh")?"zh":"en"}catch{return"zh"}}function i(t){return(De()==="zh"?Ee:Se)[t]}function K(t,r){return o=>String(o||"").toLowerCase().startsWith("zh")?t:r}function M(t){if(t==null||!Number.isFinite(t))return"";let r=["B","KB","MB","GB"],o=t,n=0;for(;o>=1024&&n<r.length-1;)o/=1024,n+=1;let s=o>=100||n===0?0:1;return`${o.toFixed(s)} ${r[n]}`}function pe(t){if(!t)return"";let r=new Date(t);if(Number.isNaN(r.getTime()))return"";try{return r.toLocaleString(void 0,{year:"numeric",month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit"})}catch{return r.toISOString()}}function W(t){let r=(t.split(".").pop()||"").toLowerCase();return["doc","docx","docm","dot","dotx","dotm","wps"].includes(r)?"\u{1F4DD}":["ppt","pptx","pptm","pps","ppsx","potx","dps"].includes(r)?"\u{1F4CA}":["xls","xlsx","xlsm","xltx","xltm","et","csv","tsv"].includes(r)?"\u{1F4C8}":"\u{1F4C4}"}var{Alert:ce,Button:D,Space:Ne,Spin:Ce,Tag:X,Tooltip:N,Typography:F}=b,{CloseOutlined:ze,DownloadOutlined:Le,ExportOutlined:Oe,FileWarningOutlined:Ie,ReloadOutlined:de}=I;function He(t,r){if(!t)return;let o=URL.createObjectURL(t),n=document.createElement("a");n.href=o,n.download=r||"document",document.body.appendChild(n),n.click(),n.remove(),window.setTimeout(()=>URL.revokeObjectURL(o),1e4)}function Me(t){let r=new Blob([t],{type:"text/html;charset=utf-8"}),o=URL.createObjectURL(r);if(!window.open(o,"_blank","noopener,noreferrer")){URL.revokeObjectURL(o);return}window.setTimeout(()=>URL.revokeObjectURL(o),6e4)}function A({state:t,onClose:r,onRetry:o,height:n="72vh",className:s}){let{status:a,result:l,filename:u,error:d}=t,c=l?.meta?.label||l?.kind?.toUpperCase()||"",g=l?.warnings??[],f=l?.meta?.converted_from,x=e.createElement("div",{className:"qopx-viewer-head"},e.createElement("div",{className:"qopx-viewer-title"},e.createElement("span",{"aria-hidden":"true"},W(u)),e.createElement("span",{title:u},u||i("routeTitle")),a==="ready"&&c&&e.createElement(X,{color:"blue"},c),f&&e.createElement(N,{title:i("convertedFrom")},e.createElement(X,null,String(f).replace(".","").toUpperCase())),l?.cached&&e.createElement(X,{color:"default"},"cache"),l?.meta?.size?e.createElement(F.Text,{type:"secondary",style:{fontSize:12}},M(l.meta.size)):null),e.createElement(Ne,{size:6,wrap:!0},a==="ready"&&l&&e.createElement(e.Fragment,null,t.blob&&e.createElement(N,{title:i("download")},e.createElement(D,{size:"small","aria-label":i("download"),icon:e.createElement(Le,null),onClick:()=>He(t.blob,u)})),e.createElement(N,{title:i("openWindow")},e.createElement(D,{size:"small","aria-label":i("openWindow"),icon:e.createElement(Oe,null),onClick:()=>Me(l.html)}))),o&&a!=="loading"&&e.createElement(N,{title:i("retry")},e.createElement(D,{size:"small","aria-label":i("retry"),icon:e.createElement(de,null),onClick:o})),r&&e.createElement(N,{title:i("close")},e.createElement(D,{size:"small",type:"text","aria-label":i("close"),icon:e.createElement(ze,null),onClick:r})))),m;return a==="loading"?m=e.createElement("div",{className:"qopx-viewer-state"},e.createElement(Ce,{size:"large"}),e.createElement(F.Text,{strong:!0},i("loading")),e.createElement(F.Text,{type:"secondary"},i("converting"))):a==="error"?m=e.createElement("div",{style:{padding:22}},e.createElement(ce,{type:"error",showIcon:!0,icon:e.createElement(Ie,null),message:i("previewFailed"),description:e.createElement("div",null,e.createElement("p",{style:{margin:"0 0 10px"}},d),t.errorCode==="legacy_converter_missing"&&e.createElement("p",{style:{margin:0,opacity:.75}},i("legacyHint")),t.errorCode==="missing_dependency"&&e.createElement("p",{style:{margin:0,opacity:.75}},i("depsMissing")),o&&e.createElement(D,{size:"small",style:{marginTop:12},icon:e.createElement(de,null),onClick:o},i("retry")))})):a==="ready"&&l?m=e.createElement(e.Fragment,null,g.length>0&&e.createElement(ce,{type:"warning",showIcon:!0,style:{borderRadius:0,borderInline:0,borderTop:0},message:i("warnings"),description:e.createElement("ul",{className:"qopx-warning-list"},g.slice(0,6).map((h,y)=>e.createElement("li",{key:`${y}-${h.slice(0,24)}`},h)))}),e.createElement("iframe",{className:"qopx-viewer-frame",title:u||i("routeTitle"),srcDoc:l.html,sandbox:"",referrerPolicy:"no-referrer"})):m=e.createElement("div",{className:"qopx-viewer-state"},e.createElement(F.Text,{type:"secondary"},i("emptyState"))),e.createElement("div",{className:`qopx-viewer${s?` ${s}`:""}`,style:{height:n}},x,e.createElement("div",{className:"qopx-viewer-body"},m))}var ue="qwenpaw-office-preview-styles",We=`
.qopx-page {
  max-width: 1180px;
  margin: 0 auto;
  padding: 22px 24px 40px;
  color: var(--ant-color-text, inherit);
  font-size: 14px;
}
.qopx-hero { margin-bottom: 18px; }
.qopx-hero h1 {
  display: flex; align-items: center; gap: 10px;
  margin: 0 0 6px; font-size: 22px; font-weight: 650;
}
.qopx-hero p { margin: 0; opacity: .72; line-height: 1.7; max-width: 760px; }
.qopx-hero-icon {
  display: inline-flex; align-items: center; justify-content: center;
  width: 36px; height: 36px; border-radius: 10px;
  background: rgba(37, 99, 235, .12);
  color: #2563eb; font-size: 18px;
}
.qopx-grid { display: grid; grid-template-columns: minmax(0, 1fr); gap: 18px; }
.qopx-card {
  background: var(--ant-color-bg-container, #fff);
  border: 1px solid var(--ant-color-border-secondary, #e5e7eb);
  border-radius: 12px;
  box-shadow: 0 1px 2px rgba(15, 23, 42, .04);
  overflow: hidden;
}
.qopx-card-head {
  display: flex; align-items: center; justify-content: space-between;
  gap: 12px; padding: 12px 16px;
  border-bottom: 1px solid var(--ant-color-border-secondary, #eef0f3);
}
.qopx-card-title { font-weight: 600; display: flex; align-items: center; gap: 8px; }
.qopx-card-body { padding: 16px; }

.qopx-drop {
  display: flex; flex-direction: column; align-items: center; justify-content: center;
  gap: 8px; width: 100%; min-height: 168px; padding: 24px;
  border: 1.5px dashed var(--ant-color-border, #cbd5e1);
  border-radius: 12px; cursor: pointer; text-align: center;
  background:
    radial-gradient(circle at 1px 1px, rgba(148, 163, 184, .18) 1px, transparent 0)
    0 0 / 16px 16px,
    var(--ant-color-fill-quaternary, #fafbfc);
  transition: border-color .18s ease, background-color .18s ease, transform .18s ease;
}
.qopx-drop:hover { border-color: #2563eb; }
.qopx-drop.qopx-dragging {
  border-color: #2563eb;
  background-color: rgba(37, 99, 235, .06);
  transform: translateY(-1px);
}
.qopx-drop-icon { font-size: 34px; line-height: 1; }
.qopx-drop-title { font-size: 15px; font-weight: 600; }
.qopx-drop-hint { font-size: 12.5px; opacity: .64; }
.qopx-drop-actions { display: flex; gap: 10px; margin-top: 6px; }

.qopx-status { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 14px; }
.qopx-chip {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 3px 10px; border-radius: 999px; font-size: 12px;
  border: 1px solid var(--ant-color-border-secondary, #e5e7eb);
  background: var(--ant-color-fill-quaternary, #fafafa);
}
.qopx-dot { width: 7px; height: 7px; border-radius: 50%; background: #22c55e; }
.qopx-dot.qopx-bad { background: #ef4444; }
.qopx-dot.qopx-muted { background: #94a3b8; }

.qopx-toolbar { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.qopx-breadcrumb {
  display: flex; align-items: center; gap: 4px; flex-wrap: wrap;
  padding: 10px 16px; border-bottom: 1px solid var(--ant-color-border-secondary, #eef0f3);
  font-size: 13px;
}
.qopx-crumb {
  border: 0; background: transparent; padding: 2px 6px; border-radius: 6px;
  cursor: pointer; color: inherit; opacity: .82;
}
.qopx-crumb:hover { background: var(--ant-color-fill-secondary, #f1f3f5); opacity: 1; }
.qopx-crumb-sep { opacity: .35; }
.qopx-file-list { max-height: 46vh; overflow: auto; }
.qopx-file-row {
  display: flex; align-items: center; gap: 10px; width: 100%;
  padding: 9px 16px; border: 0; border-bottom: 1px solid
    var(--ant-color-border-secondary, #f1f3f5);
  background: transparent; cursor: pointer; text-align: left; color: inherit;
  font-size: 13.5px;
}
.qopx-file-row:hover:not(:disabled) { background: var(--ant-color-fill-quaternary, #f8fafc); }
.qopx-file-row:disabled { cursor: not-allowed; opacity: .45; }
.qopx-file-icon {
  width: 30px; height: 30px; flex: 0 0 30px; border-radius: 8px;
  display: inline-flex; align-items: center; justify-content: center;
  font-size: 15px; background: var(--ant-color-fill-secondary, #eef1f5);
}
.qopx-file-main { min-width: 0; flex: 1; }
.qopx-file-name { display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.qopx-file-sub { display: block; font-size: 11.5px; opacity: .55; }
.qopx-file-tag { font-size: 11px; opacity: .6; white-space: nowrap; }
.qopx-empty { padding: 26px 16px; text-align: center; opacity: .6; }

.qopx-viewer { display: flex; flex-direction: column; min-height: 0; }
.qopx-viewer-head {
  display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
  padding: 10px 14px; border-bottom: 1px solid var(--ant-color-border-secondary, #eef0f3);
}
.qopx-viewer-title {
  display: flex; align-items: center; gap: 8px; min-width: 0; flex: 1;
  font-weight: 600;
}
.qopx-viewer-title span {
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.qopx-viewer-actions { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
.qopx-viewer-frame { flex: 1; min-height: 0; width: 100%; border: 0; background: #eef1f5; display: block; }
.qopx-viewer-state {
  display: flex; flex-direction: column; align-items: center; justify-content: center;
  gap: 10px; padding: 42px 20px; text-align: center; opacity: .85;
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
  background: var(--ant-color-bg-container, #fff);
  color: var(--ant-color-text, inherit);
  border-left: 1px solid var(--ant-color-border-secondary, #e5e7eb);
  box-shadow: -24px 0 56px rgba(15, 23, 42, .18),
    -2px 0 10px rgba(15, 23, 42, .06);
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
  background: color-mix(in srgb, var(--ant-color-bg-container, #fff) 92%, transparent);
  backdrop-filter: blur(14px);
}
.qopx-drawer-resizer {
  position: absolute; left: -4px; top: 0; bottom: 0; width: 9px;
  cursor: col-resize; z-index: 3; outline: none;
}
.qopx-drawer-resizer::after {
  content: ""; position: absolute; left: 3px; top: 50%;
  width: 3px; height: 42px; transform: translateY(-50%);
  border-radius: 3px; background: var(--ant-color-border, #cbd5e1);
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
  .qopx-drawer-scrim { background: rgba(15, 23, 42, .32); }
  .qopx-file-list { max-height: 52vh; }
}
`;function B(){if(typeof document>"u"||document.getElementById(ue))return;let t=document.createElement("style");t.id=ue,t.textContent=We,document.head.appendChild(t)}var{useCallback:V,useEffect:Fe,useRef:fe,useState:Ae}=e,me={status:"idle",filename:""},Be=24e4;function U(){let[t,r]=Ae(me),o=fe(0),n=fe(null),s=V(()=>{o.current+=1,n.current?.abort(),n.current=null},[]),a=V(async u=>{let d=o.current+1;o.current=d,n.current?.abort();let c=new AbortController;n.current=c,r({status:"loading",filename:u.filename});let g=window.setTimeout(()=>c.abort(),Be);try{let f=await se(u);if(d!==o.current)return;let x=f.filename||u.filename;r({status:"loading",filename:x});let m=await le(f.blob,x,c.signal);if(d!==o.current)return;r({status:"ready",filename:x,result:m,blob:f.blob})}catch(f){if(d!==o.current)return;if(f?.name==="AbortError"){r({status:"error",filename:u.filename,error:i("timeout"),errorCode:"aborted"});return}let x=f instanceof q?f:null,m=x?.status===413?i("fileTooLarge"):f?.message||i("previewFailed");r({status:"error",filename:u.filename,error:m,errorCode:x?.code})}finally{window.clearTimeout(g),n.current===c&&(n.current=null)}},[]),l=V(()=>{s(),r(me)},[s]);return Fe(()=>()=>s(),[s]),{state:t,run:a,reset:l}}var{ConfigProvider:ge}=b,G="qwenpaw-office-preview-width",Q=440,Ue=1280;function z(t){let r=Math.min(Ue,Math.max(Q,window.innerWidth-80));return Math.round(Math.min(r,Math.max(Q,t)))}function we(){return z(Math.round(window.innerWidth*.46))}function je(){try{let t=Number(window.localStorage.getItem(G));if(Number.isFinite(t)&&t>=Q)return z(t)}catch{}return we()}function _e(){let[t,r]=e.useState(je),o=e.useRef(null);return e.useEffect(()=>{let d=()=>r(c=>z(c));return window.addEventListener("resize",d),()=>window.removeEventListener("resize",d)},[]),{width:t,onPointerDown:d=>{d.preventDefault(),o.current={startX:d.clientX,startWidth:t},d.currentTarget.setPointerCapture?.(d.pointerId)},onPointerMove:d=>{if(!o.current)return;let c=o.current.startX-d.clientX;r(z(o.current.startWidth+c))},onPointerUp:d=>{if(o.current){o.current=null,d.currentTarget.releasePointerCapture?.(d.pointerId);try{window.localStorage.setItem(G,String(t))}catch{}}},resetWidth:()=>r(we()),nudgeWidth:d=>r(c=>{let g=z(c+d);try{window.localStorage.setItem(G,String(g))}catch{}return g})}}function $e({request:t,onClose:r}){let{state:o,run:n}=U(),{width:s,onPointerDown:a,onPointerMove:l,onPointerUp:u,resetWidth:d,nudgeWidth:c}=_e(),[g,f]=e.useState(!1),x=e.useRef(!1),m=e.useRef(null);e.useEffect(()=>{n(t),m.current?.focus()},[]);let h=e.useCallback(()=>{x.current||(x.current=!0,f(!0),window.setTimeout(r,180))},[r]);e.useEffect(()=>{let p=v=>{v.key==="Escape"&&(v.stopPropagation(),h())};return window.addEventListener("keydown",p,!0),()=>window.removeEventListener("keydown",p,!0)},[h]);let E=k()==="dark"&&b.theme?.darkAlgorithm?b.theme.darkAlgorithm:b.theme?.defaultAlgorithm,L=e.createElement(e.Fragment,null,e.createElement("div",{className:`qopx-drawer-scrim${g?" qopx-closing":""}`,onMouseDown:()=>h()}),e.createElement("aside",{ref:m,className:`qopx-drawer${g?" qopx-closing":""}`,style:{width:`${s}px`},role:"dialog","aria-modal":"true","aria-label":i("dialogLabel"),tabIndex:-1},e.createElement("div",{className:"qopx-drawer-resizer",role:"separator","aria-orientation":"vertical",title:i("resizeHint"),onPointerDown:a,onPointerMove:l,onPointerUp:u,onPointerCancel:u,onDoubleClick:d,tabIndex:0,"aria-label":i("resizeHint"),onKeyDown:p=>{p.key==="ArrowLeft"?(p.preventDefault(),c(32)):p.key==="ArrowRight"&&(p.preventDefault(),c(-32))}}),e.createElement(A,{state:o,onClose:h,onRetry:()=>void n(t),height:"100%"})));return ge?e.createElement(ge,{theme:E?{algorithm:E}:void 0},L):L}var C=null,j=null,xe=0,Z="",_=!1;function Re(){return C||(C=document.createElement("div"),C.id="qwenpaw-office-preview-root",document.body.appendChild(C),j=ee.createRoot(C)),j}function he(t){B();let r=Re();xe+=1,_||(Z=document.body.style.overflow,_=!0),document.body.style.overflow="hidden",r.render(e.createElement($e,{key:`preview-${xe}`,request:t,onClose:Ke}))}function Ke(){_&&(_=!1,j&&j.render(null),document.body.style.overflow=Z,Z="")}var Xe="qwenpaw:open-file-preview";function Ve(t){return t.replace(/^\/+/,"").split("/").map(r=>encodeURIComponent(r)).join("/")}function Ge(t){let r=(t.path||"").trim(),o=(t.artifactUrl||"").trim();if(!r&&!o||!P(r||o))return null;let n=T(r||o);return t.source==="workspace"&&r?{filename:n,source:{kind:"workspace",path:r,root:t.root||"project"}}:o?{filename:n,source:{kind:"artifact",url:o,path:r}}:r.startsWith("/")||/^[a-zA-Z]:[\\/]/.test(r)?{filename:n,source:{kind:"artifact",url:w.getApiUrl(`/files/preview${Ve(r)}`),path:r}}:r?{filename:n,source:{kind:"workspace",path:r,root:t.root||"project"}}:null}function Qe(t){let o=t.detail?.target;if(!o)return;let n=Ge(o);n&&(t.preventDefault(),t.stopImmediatePropagation(),he(n))}var ve=!1;function be(){ve||(ve=!0,window.addEventListener(Xe,Qe,!0))}var{Alert:ke,Button:Y,Empty:ye,Radio:Ze,Space:Ye,Spin:Je,Switch:et,Tag:tt,Typography:$}=b,{ReloadOutlined:rt,UploadOutlined:ot}=I;function nt(){let t=window.location.hash||"",r=t.indexOf("?");return r>=0?new URLSearchParams(t.slice(r+1)):new URLSearchParams(window.location.search)}function it(){let[t,r]=e.useState(null),[o,n]=e.useState(!1);if(e.useEffect(()=>{let l=!0;return ie().then(u=>{l&&r(u)}).catch(()=>{l&&n(!0)}),()=>{l=!1}},[]),o)return e.createElement($.Text,{type:"secondary",style:{fontSize:12}},i("healthFailed"));if(!t)return null;let s=[{label:"Word",ok:!!t.renderers.docx},{label:"PowerPoint",ok:!!t.renderers.pptx},{label:"Excel",ok:!!t.renderers.xlsx},{label:"LibreOffice",ok:!!t.libreoffice}],a=s.some(l=>!l.ok&&l.label!=="LibreOffice");return e.createElement(e.Fragment,null,e.createElement("div",{className:"qopx-status"},e.createElement("span",{className:"qopx-chip"},i("status")),s.map(l=>e.createElement("span",{className:"qopx-chip",key:l.label},e.createElement("span",{className:`qopx-dot ${l.ok?"":"qopx-bad"}`,"aria-hidden":"true"}),l.label,": ",l.ok?i("statusOk"):i("statusMissing")))),a&&e.createElement(ke,{type:"warning",showIcon:!0,style:{marginTop:12},message:i("depsMissing"),description:t.notes?.length?e.createElement("ul",{className:"qopx-warning-list"},t.notes.map(l=>e.createElement("li",{key:l},l))):void 0}),!t.libreoffice&&e.createElement($.Paragraph,{type:"secondary",style:{fontSize:12,margin:"10px 0 0"}},i("legacyHint")))}function at({onFile:t}){let r=e.useRef(null),[o,n]=e.useState(!1),s=a=>{let l=a&&a[0];if(l){if(!P(l.name)){b.message.warning(`${i("unsupported")} \xB7 ${i("unsupportedHint")}`);return}t(l)}};return e.createElement("div",{className:`qopx-drop${o?" qopx-dragging":""}`,role:"button",tabIndex:0,onClick:()=>r.current?.click(),onKeyDown:a=>{(a.key==="Enter"||a.key===" ")&&(a.preventDefault(),r.current?.click())},onDragOver:a=>{a.preventDefault(),n(!0)},onDragLeave:a=>{a.preventDefault(),n(!1)},onDrop:a=>{a.preventDefault(),n(!1),s(a.dataTransfer?.files??null)}},e.createElement("span",{className:"qopx-drop-icon","aria-hidden":"true"},"\u{1F4C4}"),e.createElement("span",{className:"qopx-drop-title"},i("dropTitle")),e.createElement("span",{className:"qopx-drop-hint"},i("dropHint")),e.createElement("div",{className:"qopx-drop-actions"},e.createElement(Y,{type:"primary",size:"small",icon:e.createElement(ot,null),onClick:a=>{a.stopPropagation(),r.current?.click()}},i("chooseFile"))),e.createElement("input",{ref:r,type:"file",accept:ne,style:{display:"none"},onChange:a=>{s(a.target.files),a.target.value=""}}))}function st({onPick:t}){let[r,o]=e.useState("project"),[n,s]=e.useState(""),[a,l]=e.useState(null),[u,d]=e.useState(!1),[c,g]=e.useState(""),[f,x]=e.useState(!0),m=e.useRef(0),h=e.useCallback(async(p,v)=>{let O=m.current+1;m.current=O,d(!0),g("");try{let R=await ae(p,v);O===m.current&&l(R)}catch(R){O===m.current&&(l(null),g(R?.message||i("loadDirFailed")))}finally{O===m.current&&d(!1)}},[]);e.useEffect(()=>{h(n,r)},[n,r,h]);let y=n.split("/").filter(Boolean),E=(a?.entries??[]).filter(p=>p.kind==="directory"?!0:!f||P(p.path)),L=()=>s(y.slice(0,-1).join("/"));return e.createElement(e.Fragment,null,e.createElement("div",{className:"qopx-breadcrumb"},e.createElement(Ze.Group,{size:"small",value:r,optionType:"button",options:[{label:i("projectRoot"),value:"project"},{label:i("workspaceRoot"),value:"workspace"}],onChange:p=>{s(""),o(p.target.value)}}),e.createElement("span",{style:{width:8}}),e.createElement("button",{type:"button",className:"qopx-crumb",onClick:()=>s(""),title:i("root")},r==="workspace"?i("workspaceRoot"):i("projectRoot")),y.map((p,v)=>e.createElement(e.Fragment,{key:`${p}-${v}`},e.createElement("span",{className:"qopx-crumb-sep"},"/"),e.createElement("button",{type:"button",className:"qopx-crumb",onClick:()=>s(y.slice(0,v+1).join("/"))},p))),e.createElement("span",{style:{flex:1}}),e.createElement(Ye,{size:6},n&&e.createElement(Y,{size:"small",type:"text",onClick:L},i("parent")),e.createElement(Y,{size:"small",type:"text",icon:e.createElement(rt,null),onClick:()=>void h(n,r)},i("refresh")),e.createElement($.Text,{type:"secondary",style:{fontSize:12}},i("filterOnly")),e.createElement(et,{size:"small",checked:f,onChange:x}))),c&&e.createElement("div",{style:{padding:16}},e.createElement(ke,{type:"error",showIcon:!0,message:i("loadDirFailed"),description:c})),e.createElement("div",{className:"qopx-file-list"},u?e.createElement("div",{className:"qopx-empty"},e.createElement(Je,{size:"small"})," ",e.createElement("span",{style:{marginLeft:8}},i("loading"))):E.length===0?e.createElement("div",{className:"qopx-empty"},e.createElement(ye,{image:ye.PRESENTED_IMAGE_SIMPLE,description:i("emptyDir")})):E.map(p=>{let v=p.kind==="directory"||P(p.path);return e.createElement("button",{type:"button",key:p.path,className:"qopx-file-row",disabled:!v,onClick:()=>{if(p.kind==="directory"){s(p.path);return}t({filename:p.name,source:{kind:"workspace",path:p.path,root:r}})}},e.createElement("span",{className:"qopx-file-icon","aria-hidden":"true"},p.kind==="directory"?"\u{1F4C1}":W(p.name)),e.createElement("span",{className:"qopx-file-main"},e.createElement("span",{className:"qopx-file-name",title:p.name},p.name),e.createElement("span",{className:"qopx-file-sub"},p.kind==="directory"?"folder":[M(p.size),pe(p.modified_at)].filter(Boolean).join(" \xB7 "))),p.kind==="file"&&!v&&e.createElement(tt,null,i("unsupported")))})))}function qe(){let{state:t,run:r,reset:o}=U(),[n,s]=e.useState(null),a=e.useRef(null),l=e.useRef(!1),u=e.useCallback(c=>{s(c),r(c),window.setTimeout(()=>{a.current?.scrollIntoView({behavior:"smooth",block:"start"})},30)},[r]);e.useEffect(()=>{if(l.current)return;l.current=!0;let c=nt(),g=c.get("artifact"),f=c.get("path"),x=c.get("name")||void 0;g?u({filename:x||T(g),source:{kind:"artifact",url:g}}):f&&u({filename:x||T(f),source:{kind:"workspace",path:f,root:c.get("root")||"project"}})},[u]);let d=null;return t.status!=="idle"&&(d=e.createElement("section",{className:"qopx-card",ref:a},e.createElement(A,{state:t,onClose:o,onRetry:n?()=>void r(n):void 0,height:"70vh"}))),e.createElement("div",{className:"qopx-page","data-theme":k()},e.createElement("header",{className:"qopx-hero"},e.createElement("h1",null,e.createElement("span",{className:"qopx-hero-icon","aria-hidden":"true"},"\u{1F4C4}"),i("routeTitle")),e.createElement("p",null,i("pageSubtitle"))),e.createElement("div",{className:"qopx-grid"},e.createElement("section",{className:"qopx-card"},e.createElement("div",{className:"qopx-card-body"},e.createElement(at,{onFile:c=>u({filename:c.name,source:{kind:"upload",blob:c}})}),e.createElement(it,null))),d,e.createElement("section",{className:"qopx-card"},e.createElement("div",{className:"qopx-card-head"},e.createElement("div",{className:"qopx-card-title"},"\u{1F5C2} ",i("workspace")),e.createElement($.Text,{type:"secondary",style:{fontSize:12}},i("cacheNote"))),e.createElement(st,{onPick:u}))))}var J="qwenpaw-office-preview";B();be();try{window.QwenPaw.registerRoutes?.(J,[{path:`/plugin/${J}`,component:qe,label:K("Office \u9884\u89C8","Office Preview"),title:K("Office \u6587\u6863\u9884\u89C8","Office Document Preview"),icon:"\u{1F4C4}",priority:46}])}catch(t){console.error(`[${J}] failed to register route`,t)}try{window.QwenPawOfficePreview={open:t=>{let r={source:t.artifact?"attachment":"workspace",path:t.path||"",artifactUrl:t.artifact,root:t.root||"project"};window.dispatchEvent(new CustomEvent("qwenpaw:open-file-preview",{detail:{target:r}}))},title:i("routeLabel")}}catch{}
