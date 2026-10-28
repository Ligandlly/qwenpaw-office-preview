var v=window.QwenPaw.host,e=v.React,te=v.ReactDOM,y=v.antd,T=v.antdIcons;function I(){try{if(document.documentElement.classList.contains("dark-mode"))return"dark";v.useTheme}catch{}return"light"}function H(){let[t,r]=e.useState(I);return e.useEffect(()=>{let o=()=>r(I()),n=new MutationObserver(o);return n.observe(document.documentElement,{attributes:!0,attributeFilter:["class","data-theme"]}),window.addEventListener("qwenpaw:auth-changed",o),()=>{n.disconnect(),window.removeEventListener("qwenpaw:auth-changed",o)}},[]),t}function Ce(){try{let t=window.sessionStorage.getItem("qwenpaw-agent-storage")??window.localStorage.getItem("qwenpaw-agent-storage");if(t){let o=JSON.parse(t)?.state?.selectedAgent;if(typeof o=="string"&&o)return o}}catch{}try{return v.getSelectedAgentId?.()??"default"}catch{return"default"}}function re(){let t={};try{let o=v.getApiToken();o&&(t.Authorization=`Bearer ${o}`)}catch{}let r=Ce();return r&&(t["X-Agent-Id"]=r),t}async function C(t,r={}){if(v.fetch)return v.fetch(t,r);let o=/^https?:\/\//i.test(t)?t:v.getApiUrl(t);return fetch(o,{...r,headers:{...re(),...r.headers??{}}})}async function oe(t,r={}){return fetch(t,{...r,headers:{...re(),...r.headers??{}}})}var ne=new Set(["csv","doc","docm","docx","dot","dotm","dotx","dps","et","pps","ppsx","ppt","pptm","pptx","potx","tsv","wps","xls","xlsm","xlsx","xltm","xltx"]),ie=Array.from(ne).map(t=>`.${t}`).join(","),E=class extends Error{constructor(r,o,n){super(r),this.name="ApiError",this.status=o,this.code=n}};function Ne(t){let o=(t||"").split(/[?#]/,1)[0].replace(/\\/g,"/").split("/").pop()||"",n=o.lastIndexOf(".");return n>=0?o.slice(n+1).toLowerCase():""}function S(t){return ne.has(Ne(t))}function D(t){return(t||"").replace(/\\/g,"/").split(/[?#]/,1)[0].split("/").filter(Boolean).pop()||"document"}async function F(t){let r=t.status,o=`${r} ${t.statusText}`.trim(),n=`http_${r}`;try{let s=await t.json(),a=s?.detail??s;typeof a=="string"?o=a:a&&typeof a=="object"&&(typeof a.message=="string"&&(o=a.message),typeof a.code=="string"&&(n=a.code))}catch{try{let s=await t.text();s&&(o=s.slice(0,500))}catch{}}return new E(o,r,n)}async function ae(){let t=await C("/office-preview/health");if(!t.ok)throw await F(t);return await t.json()}async function se(t,r,o=400){let n=new URLSearchParams({path:t,root:r,limit:String(o)}),s=await C(`/workspace/tree?${n.toString()}`);if(!s.ok)throw await F(s);return await s.json()}async function le(t){let{source:r}=t;if(r.kind==="upload")return{blob:r.blob,filename:t.filename};let o;if(r.kind==="workspace"){let a=new URLSearchParams({path:r.path,root:r.root||"project"});o=await C(`/workspace/file-download?${a.toString()}`)}else o=await oe(r.url);if(!o.ok)throw await F(o);let n=await o.blob(),s=t.filename||D(r.kind==="workspace"?r.path:r.path||"");return{blob:n,filename:s}}async function pe(t,r,o){let n=new URLSearchParams({filename:r||"document.docx",theme:I()}),s=await C(`/office-preview/convert?${n.toString()}`,{method:"POST",headers:{"Content-Type":"application/octet-stream"},body:t,signal:o});if(!s.ok)throw await F(s);let a=await s.json();if(!a||typeof a.html!="string"||!a.html)throw new E("The server returned an empty preview.",500,"empty_result");return a}var Oe={routeLabel:"Office \u9884\u89C8",routeTitle:"Office \u6587\u6863\u9884\u89C8",pageSubtitle:"\u65E0\u9700\u4E0B\u8F7D\u5373\u53EF\u67E5\u770B Word\u3001PowerPoint \u548C Excel \u6587\u6863\u3002\u62D6\u5165\u672C\u5730\u6587\u4EF6\uFF0C\u6216\u76F4\u63A5\u4ECE\u5F53\u524D\u5DE5\u4F5C\u533A\u6253\u5F00\u3002",dropTitle:"\u62D6\u62FD\u6587\u4EF6\u5230\u8FD9\u91CC\uFF0C\u6216\u70B9\u51FB\u9009\u62E9",dropHint:"\u652F\u6301 doc / docx / ppt / pptx / xls / xlsx\uFF0C\u4EE5\u53CA csv / tsv",chooseFile:"\u9009\u62E9\u6587\u4EF6",workspace:"\u5DE5\u4F5C\u533A\u6587\u4EF6",projectRoot:"\u9879\u76EE\u76EE\u5F55",workspaceRoot:"\u5DE5\u4F5C\u533A\u76EE\u5F55",refresh:"\u5237\u65B0",filterOnly:"\u4EC5\u663E\u793A\u53EF\u9884\u89C8",emptyDir:"\u8FD9\u4E2A\u76EE\u5F55\u91CC\u6CA1\u6709\u53EF\u9884\u89C8\u7684\u6587\u6863\u3002",loadDirFailed:"\u8BFB\u53D6\u76EE\u5F55\u5931\u8D25",unsupported:"\u4E0D\u652F\u6301\u9884\u89C8\u8BE5\u6587\u4EF6\u7C7B\u578B",unsupportedHint:"\u652F\u6301 Word\u3001PowerPoint\u3001Excel \u6587\u6863\u4EE5\u53CA CSV\u3002",loading:"\u6B63\u5728\u751F\u6210\u9884\u89C8\u2026",converting:"\u6B63\u5728\u89E3\u6790\u6587\u6863\uFF0C\u8F83\u5927\u7684\u6587\u4EF6\u53EF\u80FD\u9700\u8981\u51E0\u79D2\u949F\u3002",retry:"\u91CD\u8BD5",close:"\u5173\u95ED",download:"\u4E0B\u8F7D\u539F\u6587\u4EF6",openWindow:"\u65B0\u7A97\u53E3\u6253\u5F00",previewFailed:"\u9884\u89C8\u5931\u8D25",backToList:"\u8FD4\u56DE\u6587\u4EF6\u5217\u8868",size:"\u5927\u5C0F",modified:"\u4FEE\u6539\u65F6\u95F4",status:"\u8F6C\u6362\u5668\u72B6\u6001",statusOk:"\u53EF\u7528",statusMissing:"\u672A\u5B89\u88C5",healthFailed:"\u65E0\u6CD5\u83B7\u53D6\u8F6C\u6362\u5668\u72B6\u6001",depsMissing:"\u90E8\u5206\u8F6C\u6362\u5668\u4F9D\u8D56\u7F3A\u5931\uFF0C\u8BF7\u91CD\u65B0\u5B89\u88C5\u63D2\u4EF6\u4EE5\u81EA\u52A8\u8865\u9F50\u4F9D\u8D56\u3002",legacyHint:"\u65E7\u7248 .doc / .ppt / .xls \u6587\u4EF6\u9700\u8981\u672C\u673A\u5B89\u88C5 LibreOffice\uFF1B\u4E5F\u53EF\u4EE5\u5148\u53E6\u5B58\u4E3A docx / pptx / xlsx\u3002",cacheNote:"\u8F6C\u6362\u7ED3\u679C\u4F1A\u7F13\u5B58\u5728\u5185\u5B58\u4E2D\uFF0C\u91CD\u590D\u6253\u5F00\u540C\u4E00\u6587\u4EF6\u4F1A\u66F4\u5FEB\u3002",fileTooLarge:"\u6587\u4EF6\u8FC7\u5927\uFF0C\u65E0\u6CD5\u5728\u7EBF\u9884\u89C8\u3002",dialogLabel:"Office \u6587\u6863\u9884\u89C8",navigate:"\u5B9A\u4F4D",parent:"\u4E0A\u4E00\u7EA7",root:"\u6839\u76EE\u5F55",entries:"\u9879",openInPage:"\u5728\u72EC\u7ACB\u9875\u9762\u4E2D\u6253\u5F00",emptyState:"\u9009\u62E9\u5DE6\u4FA7\u6587\u4EF6\u6216\u62D6\u5165\u6587\u6863\u5F00\u59CB\u9884\u89C8",convertedFrom:"\u5DF2\u4ECE\u65E7\u683C\u5F0F\u8F6C\u6362",warnings:"\u683C\u5F0F\u63D0\u793A",timeout:"\u9884\u89C8\u5DF2\u53D6\u6D88\u6216\u8D85\u65F6\u3002",resizeHint:"\u62D6\u52A8\u8C03\u6574\u5BBD\u5EA6\uFF0C\u53CC\u51FB\u6062\u590D\u9ED8\u8BA4"},Le={routeLabel:"Office Preview",routeTitle:"Office Document Preview",pageSubtitle:"Read Word, PowerPoint and Excel documents without downloading them. Drop a local file or open one from the workspace.",dropTitle:"Drop a file here, or click to choose",dropHint:"Supports doc / docx / ppt / pptx / xls / xlsx, plus csv / tsv",chooseFile:"Choose file",workspace:"Workspace files",projectRoot:"Project",workspaceRoot:"Workspace",refresh:"Refresh",filterOnly:"Previewable only",emptyDir:"No previewable documents in this folder.",loadDirFailed:"Failed to read the folder",unsupported:"This file type cannot be previewed",unsupportedHint:"Word, PowerPoint, Excel and CSV documents are supported.",loading:"Building preview\u2026",converting:"Parsing the document; large files may take a few seconds.",retry:"Retry",close:"Close",download:"Download original",openWindow:"Open in new window",previewFailed:"Preview failed",backToList:"Back to files",size:"Size",modified:"Modified",status:"Converter status",statusOk:"Available",statusMissing:"Not installed",healthFailed:"Could not read converter status",depsMissing:"Some converters are unavailable. Reinstall the plugin to restore its Python dependencies.",legacyHint:"Legacy .doc / .ppt / .xls previews require LibreOffice on the host; alternatively save the file as docx / pptx / xlsx.",cacheNote:"Converted documents are cached in memory, so reopening the same file is faster.",fileTooLarge:"The file is too large to preview online.",dialogLabel:"Office document preview",navigate:"Go",parent:"Up",root:"Root",entries:"items",openInPage:"Open in dedicated page",emptyState:"Pick a file on the left or drop a document to begin",convertedFrom:"Converted from a legacy format",warnings:"Document notes",timeout:"Preview was cancelled or timed out.",resizeHint:"Drag to resize, double-click to reset"};function ze(){try{return(window.localStorage.getItem("language")||document.documentElement.getAttribute("lang")||navigator.language||"en").toLowerCase().startsWith("zh")?"zh":"en"}catch{return"zh"}}function i(t){return(ze()==="zh"?Oe:Le)[t]}function K(t,r){return o=>String(o||"").toLowerCase().startsWith("zh")?t:r}function M(t){if(t==null||!Number.isFinite(t))return"";let r=["B","KB","MB","GB"],o=t,n=0;for(;o>=1024&&n<r.length-1;)o/=1024,n+=1;let s=o>=100||n===0?0:1;return`${o.toFixed(s)} ${r[n]}`}function ce(t){if(!t)return"";let r=new Date(t);if(Number.isNaN(r.getTime()))return"";try{return r.toLocaleString(void 0,{year:"numeric",month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit"})}catch{return r.toISOString()}}var{CloudUploadOutlined:Ie,FileExcelOutlined:He,FilePptOutlined:Fe,FileTextOutlined:de,FileWordOutlined:Me,FolderOpenOutlined:We,FolderOutlined:Ae}=T;function Ue(t){let r=(t||"").replace(/\\/g,"/").split("/").pop()||"",o=r.lastIndexOf(".");return o>=0?r.slice(o+1).toLowerCase():""}var Be=new Set(["doc","docm","docx","dot","dotm","dotx","wps"]),je=new Set(["dps","potx","pps","ppsx","ppt","pptm","pptx"]),_e=new Set(["csv","et","tsv","xls","xlsm","xlsx","xltm","xltx"]);function W({name:t,directory:r=!1,style:o}){if(r)return e.createElement(Ae,{style:o});let n=Ue(t);return Be.has(n)?e.createElement(Me,{style:o}):je.has(n)?e.createElement(Fe,{style:o}):_e.has(n)?e.createElement(He,{style:o}):e.createElement(de,{style:o})}function ue(){return e.createElement(Ie,null)}function fe(){return e.createElement(de,null)}function me(){return e.createElement(We,null)}var{Alert:xe,Button:N,Space:$e,Spin:Re,Tag:X,Tooltip:O,Typography:A}=y,{CloseOutlined:Ke,DownloadOutlined:Xe,ExportOutlined:Ve,FileWarningOutlined:Ge,ReloadOutlined:ge}=T;function Qe(t,r){if(!t)return;let o=URL.createObjectURL(t),n=document.createElement("a");n.href=o,n.download=r||"document",document.body.appendChild(n),n.click(),n.remove(),window.setTimeout(()=>URL.revokeObjectURL(o),1e4)}function Ze(t){let r=new Blob([t],{type:"text/html;charset=utf-8"}),o=URL.createObjectURL(r);if(!window.open(o,"_blank","noopener,noreferrer")){URL.revokeObjectURL(o);return}window.setTimeout(()=>URL.revokeObjectURL(o),6e4)}function U({state:t,onClose:r,onRetry:o,height:n="72vh",className:s}){let{status:a,result:l,filename:f,error:p}=t,u=l?.meta?.label||l?.kind?.toUpperCase()||"",g=l?.warnings??[],c=l?.meta?.converted_from,x=e.createElement("div",{className:"qopx-viewer-head"},e.createElement("div",{className:"qopx-viewer-title"},e.createElement("span",{className:"qopx-viewer-file-icon","aria-hidden":"true"},e.createElement(W,{name:f})),e.createElement("span",{title:f},f||i("routeTitle")),a==="ready"&&u&&e.createElement(X,{color:"blue"},u),c&&e.createElement(O,{title:i("convertedFrom")},e.createElement(X,null,String(c).replace(".","").toUpperCase())),l?.cached&&e.createElement(X,{color:"default"},"cache"),l?.meta?.size?e.createElement(A.Text,{type:"secondary",style:{fontSize:12}},M(l.meta.size)):null),e.createElement($e,{size:6,wrap:!0},a==="ready"&&l&&e.createElement(e.Fragment,null,t.blob&&e.createElement(O,{title:i("download")},e.createElement(N,{size:"small","aria-label":i("download"),icon:e.createElement(Xe,null),onClick:()=>Qe(t.blob,f)})),e.createElement(O,{title:i("openWindow")},e.createElement(N,{size:"small","aria-label":i("openWindow"),icon:e.createElement(Ve,null),onClick:()=>Ze(l.html)}))),o&&a!=="loading"&&e.createElement(O,{title:i("retry")},e.createElement(N,{size:"small","aria-label":i("retry"),icon:e.createElement(ge,null),onClick:o})),r&&e.createElement(O,{title:i("close")},e.createElement(N,{size:"small",type:"text","aria-label":i("close"),icon:e.createElement(Ke,null),onClick:r})))),m;return a==="loading"?m=e.createElement("div",{className:"qopx-viewer-state"},e.createElement(Re,{size:"large"}),e.createElement(A.Text,{strong:!0},i("loading")),e.createElement(A.Text,{type:"secondary"},i("converting"))):a==="error"?m=e.createElement("div",{style:{padding:22}},e.createElement(xe,{type:"error",showIcon:!0,icon:e.createElement(Ge,null),message:i("previewFailed"),description:e.createElement("div",null,e.createElement("p",{style:{margin:"0 0 10px"}},p),t.errorCode==="legacy_converter_missing"&&e.createElement("p",{style:{margin:0,opacity:.75}},i("legacyHint")),t.errorCode==="missing_dependency"&&e.createElement("p",{style:{margin:0,opacity:.75}},i("depsMissing")),o&&e.createElement(N,{size:"small",style:{marginTop:12},icon:e.createElement(ge,null),onClick:o},i("retry")))})):a==="ready"&&l?m=e.createElement(e.Fragment,null,g.length>0&&e.createElement(xe,{type:"warning",showIcon:!0,style:{borderRadius:0,borderInline:0,borderTop:0},message:i("warnings"),description:e.createElement("ul",{className:"qopx-warning-list"},g.slice(0,6).map((w,q)=>e.createElement("li",{key:`${q}-${w.slice(0,24)}`},w)))}),e.createElement("iframe",{className:"qopx-viewer-frame",title:f||i("routeTitle"),srcDoc:l.html,sandbox:"",referrerPolicy:"no-referrer"})):m=e.createElement("div",{className:"qopx-viewer-state"},e.createElement(A.Text,{type:"secondary"},i("emptyState"))),e.createElement("div",{className:`qopx-viewer${s?` ${s}`:""}`,style:{height:n}},x,e.createElement("div",{className:"qopx-viewer-body"},m))}var we="qwenpaw-office-preview-styles",Ye=`
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
`;function B(){if(typeof document>"u"||document.getElementById(we))return;let t=document.createElement("style");t.id=we,t.textContent=Ye,document.head.appendChild(t)}var{useCallback:V,useEffect:Je,useRef:he,useState:et}=e,ve={status:"idle",filename:""},tt=24e4;function j(){let[t,r]=et(ve),o=he(0),n=he(null),s=V(()=>{o.current+=1,n.current?.abort(),n.current=null},[]),a=V(async f=>{let p=o.current+1;o.current=p,n.current?.abort();let u=new AbortController;n.current=u,r({status:"loading",filename:f.filename});let g=window.setTimeout(()=>u.abort(),tt);try{let c=await le(f);if(p!==o.current)return;let x=c.filename||f.filename;r({status:"loading",filename:x});let m=await pe(c.blob,x,u.signal);if(p!==o.current)return;r({status:"ready",filename:x,result:m,blob:c.blob})}catch(c){if(p!==o.current)return;if(c?.name==="AbortError"){r({status:"error",filename:f.filename,error:i("timeout"),errorCode:"aborted"});return}let x=c instanceof E?c:null,m=x?.status===413?i("fileTooLarge"):c?.message||i("previewFailed");r({status:"error",filename:f.filename,error:m,errorCode:x?.code})}finally{window.clearTimeout(g),n.current===u&&(n.current=null)}},[]),l=V(()=>{s(),r(ve)},[s]);return Je(()=>()=>s(),[s]),{state:t,run:a,reset:l}}var{ConfigProvider:be}=y,G="qwenpaw-office-preview-width",Q=440,rt=1280;function z(t){let r=Math.min(rt,Math.max(Q,window.innerWidth-80));return Math.round(Math.min(r,Math.max(Q,t)))}function qe(){return z(Math.round(window.innerWidth*.46))}function ot(){try{let t=Number(window.localStorage.getItem(G));if(Number.isFinite(t)&&t>=Q)return z(t)}catch{}return qe()}function nt(){let[t,r]=e.useState(ot),o=e.useRef(null);return e.useEffect(()=>{let p=()=>r(u=>z(u));return window.addEventListener("resize",p),()=>window.removeEventListener("resize",p)},[]),{width:t,onPointerDown:p=>{p.preventDefault(),o.current={startX:p.clientX,startWidth:t},p.currentTarget.setPointerCapture?.(p.pointerId)},onPointerMove:p=>{if(!o.current)return;let u=o.current.startX-p.clientX;r(z(o.current.startWidth+u))},onPointerUp:p=>{if(o.current){o.current=null,p.currentTarget.releasePointerCapture?.(p.pointerId);try{window.localStorage.setItem(G,String(t))}catch{}}},resetWidth:()=>r(qe()),nudgeWidth:p=>r(u=>{let g=z(u+p);try{window.localStorage.setItem(G,String(g))}catch{}return g})}}function it({request:t,onClose:r}){let{state:o,run:n}=j(),s=H(),a=e.useRef(s),{width:l,onPointerDown:f,onPointerMove:p,onPointerUp:u,resetWidth:g,nudgeWidth:c}=nt(),[x,m]=e.useState(!1),w=e.useRef(!1),q=e.useRef(null);e.useEffect(()=>{n(t),q.current?.focus()},[]),e.useEffect(()=>{a.current!==s&&(a.current=s,n(t))},[s,t,n]);let k=e.useCallback(()=>{w.current||(w.current=!0,m(!0),window.setTimeout(r,180))},[r]);e.useEffect(()=>{let h=P=>{P.key==="Escape"&&(P.stopPropagation(),k())};return window.addEventListener("keydown",h,!0),()=>window.removeEventListener("keydown",h,!0)},[k]);let d=s==="dark"&&y.theme?.darkAlgorithm?y.theme.darkAlgorithm:y.theme?.defaultAlgorithm,b=e.createElement(e.Fragment,null,e.createElement("div",{className:`qopx-drawer-scrim${x?" qopx-closing":""}`,onMouseDown:()=>k()}),e.createElement("aside",{ref:q,className:`qopx-drawer${x?" qopx-closing":""}`,style:{width:`${l}px`},role:"dialog","aria-modal":"true","aria-label":i("dialogLabel"),"data-theme":s,tabIndex:-1},e.createElement("div",{className:"qopx-drawer-resizer",role:"separator","aria-orientation":"vertical",title:i("resizeHint"),onPointerDown:f,onPointerMove:p,onPointerUp:u,onPointerCancel:u,onDoubleClick:g,tabIndex:0,"aria-label":i("resizeHint"),onKeyDown:h=>{h.key==="ArrowLeft"?(h.preventDefault(),c(32)):h.key==="ArrowRight"&&(h.preventDefault(),c(-32))}}),e.createElement(U,{state:o,onClose:k,onRetry:()=>void n(t),height:"100%"})));return be?e.createElement(be,{theme:d?{algorithm:d}:void 0},b):b}var L=null,_=null,ye=0,Z="",$=!1;function at(){return L||(L=document.createElement("div"),L.id="qwenpaw-office-preview-root",document.body.appendChild(L),_=te.createRoot(L)),_}function ke(t){B();let r=at();ye+=1,$||(Z=document.body.style.overflow,$=!0),document.body.style.overflow="hidden",r.render(e.createElement(it,{key:`preview-${ye}`,request:t,onClose:st}))}function st(){$&&($=!1,_&&_.render(null),document.body.style.overflow=Z,Z="")}var lt="qwenpaw:open-file-preview";function pt(t){return t.replace(/^\/+/,"").split("/").map(r=>encodeURIComponent(r)).join("/")}function ct(t){let r=(t.path||"").trim(),o=(t.artifactUrl||"").trim();if(!r&&!o||!S(r||o))return null;let n=D(r||o);return t.source==="workspace"&&r?{filename:n,source:{kind:"workspace",path:r,root:t.root||"project"}}:o?{filename:n,source:{kind:"artifact",url:o,path:r}}:r.startsWith("/")||/^[a-zA-Z]:[\\/]/.test(r)?{filename:n,source:{kind:"artifact",url:v.getApiUrl(`/files/preview${pt(r)}`),path:r}}:r?{filename:n,source:{kind:"workspace",path:r,root:t.root||"project"}}:null}function dt(t){let o=t.detail?.target;if(!o)return;let n=ct(o);n&&(t.preventDefault(),t.stopImmediatePropagation(),ke(n))}var Pe=!1;function Te(){Pe||(Pe=!0,window.addEventListener(lt,dt,!0))}var{Alert:Se,Button:Y,Empty:Ee,Radio:ut,Space:ft,Spin:mt,Switch:xt,Tag:gt,Typography:R}=y,{ReloadOutlined:wt,UploadOutlined:ht}=T;function vt(){let t=window.location.hash||"",r=t.indexOf("?");return r>=0?new URLSearchParams(t.slice(r+1)):new URLSearchParams(window.location.search)}function bt(){let[t,r]=e.useState(null),[o,n]=e.useState(!1);if(e.useEffect(()=>{let l=!0;return ae().then(f=>{l&&r(f)}).catch(()=>{l&&n(!0)}),()=>{l=!1}},[]),o)return e.createElement(R.Text,{type:"secondary",style:{fontSize:12}},i("healthFailed"));if(!t)return null;let s=[{label:"Word",ok:!!t.renderers.docx},{label:"PowerPoint",ok:!!t.renderers.pptx},{label:"Excel",ok:!!t.renderers.xlsx},{label:"LibreOffice",ok:!!t.libreoffice}],a=s.some(l=>!l.ok&&l.label!=="LibreOffice");return e.createElement(e.Fragment,null,e.createElement("div",{className:"qopx-status"},e.createElement("span",{className:"qopx-chip"},i("status")),s.map(l=>e.createElement("span",{className:"qopx-chip",key:l.label},e.createElement("span",{className:`qopx-dot ${l.ok?"":"qopx-bad"}`,"aria-hidden":"true"}),l.label,": ",l.ok?i("statusOk"):i("statusMissing")))),a&&e.createElement(Se,{type:"warning",showIcon:!0,style:{marginTop:12},message:i("depsMissing"),description:t.notes?.length?e.createElement("ul",{className:"qopx-warning-list"},t.notes.map(l=>e.createElement("li",{key:l},l))):void 0}),!t.libreoffice&&e.createElement(R.Paragraph,{type:"secondary",style:{fontSize:12,margin:"10px 0 0"}},i("legacyHint")))}function yt({onFile:t}){let r=e.useRef(null),[o,n]=e.useState(!1),s=a=>{let l=a&&a[0];if(l){if(!S(l.name)){y.message.warning(`${i("unsupported")} \xB7 ${i("unsupportedHint")}`);return}t(l)}};return e.createElement("div",{className:`qopx-drop${o?" qopx-dragging":""}`,role:"button",tabIndex:0,onClick:()=>r.current?.click(),onKeyDown:a=>{(a.key==="Enter"||a.key===" ")&&(a.preventDefault(),r.current?.click())},onDragOver:a=>{a.preventDefault(),n(!0)},onDragLeave:a=>{a.preventDefault(),n(!1)},onDrop:a=>{a.preventDefault(),n(!1),s(a.dataTransfer?.files??null)}},e.createElement("span",{className:"qopx-drop-icon","aria-hidden":"true"},e.createElement(ue,null)),e.createElement("span",{className:"qopx-drop-title"},i("dropTitle")),e.createElement("span",{className:"qopx-drop-hint"},i("dropHint")),e.createElement("div",{className:"qopx-drop-actions"},e.createElement(Y,{type:"primary",size:"small",icon:e.createElement(ht,null),onClick:a=>{a.stopPropagation(),r.current?.click()}},i("chooseFile"))),e.createElement("input",{ref:r,type:"file",accept:ie,style:{display:"none"},onChange:a=>{s(a.target.files),a.target.value=""}}))}function qt({onPick:t}){let[r,o]=e.useState("project"),[n,s]=e.useState(""),[a,l]=e.useState(null),[f,p]=e.useState(!1),[u,g]=e.useState(""),[c,x]=e.useState(!0),m=e.useRef(0),w=e.useCallback(async(d,b)=>{let h=m.current+1;m.current=h,p(!0),g("");try{let P=await se(d,b);h===m.current&&l(P)}catch(P){h===m.current&&(l(null),g(P?.message||i("loadDirFailed")))}finally{h===m.current&&p(!1)}},[]);e.useEffect(()=>{w(n,r)},[n,r,w]);let q=n.split("/").filter(Boolean),k=(a?.entries??[]).filter(d=>d.kind==="directory"?!0:!c||S(d.path)),ee=()=>s(q.slice(0,-1).join("/"));return e.createElement(e.Fragment,null,e.createElement("div",{className:"qopx-breadcrumb"},e.createElement(ut.Group,{size:"small",value:r,optionType:"button",options:[{label:i("projectRoot"),value:"project"},{label:i("workspaceRoot"),value:"workspace"}],onChange:d=>{s(""),o(d.target.value)}}),e.createElement("span",{style:{width:8}}),e.createElement("button",{type:"button",className:"qopx-crumb",onClick:()=>s(""),title:i("root")},r==="workspace"?i("workspaceRoot"):i("projectRoot")),q.map((d,b)=>e.createElement(e.Fragment,{key:`${d}-${b}`},e.createElement("span",{className:"qopx-crumb-sep"},"/"),e.createElement("button",{type:"button",className:"qopx-crumb",onClick:()=>s(q.slice(0,b+1).join("/"))},d))),e.createElement("span",{style:{flex:1}}),e.createElement(ft,{size:6},n&&e.createElement(Y,{size:"small",type:"text",onClick:ee},i("parent")),e.createElement(Y,{size:"small",type:"text",icon:e.createElement(wt,null),onClick:()=>void w(n,r)},i("refresh")),e.createElement(R.Text,{type:"secondary",style:{fontSize:12}},i("filterOnly")),e.createElement(xt,{size:"small",checked:c,onChange:x}))),u&&e.createElement("div",{style:{padding:16}},e.createElement(Se,{type:"error",showIcon:!0,message:i("loadDirFailed"),description:u})),e.createElement("div",{className:"qopx-file-list"},f?e.createElement("div",{className:"qopx-empty"},e.createElement(mt,{size:"small"})," ",e.createElement("span",{style:{marginLeft:8}},i("loading"))):k.length===0?e.createElement("div",{className:"qopx-empty"},e.createElement(Ee,{image:Ee.PRESENTED_IMAGE_SIMPLE,description:i("emptyDir")})):k.map(d=>{let b=d.kind==="directory"||S(d.path);return e.createElement("button",{type:"button",key:d.path,className:"qopx-file-row",disabled:!b,onClick:()=>{if(d.kind==="directory"){s(d.path);return}t({filename:d.name,source:{kind:"workspace",path:d.path,root:r}})}},e.createElement("span",{className:"qopx-file-icon","aria-hidden":"true"},e.createElement(W,{name:d.name,directory:d.kind==="directory"})),e.createElement("span",{className:"qopx-file-main"},e.createElement("span",{className:"qopx-file-name",title:d.name},d.name),e.createElement("span",{className:"qopx-file-sub"},d.kind==="directory"?"folder":[M(d.size),ce(d.modified_at)].filter(Boolean).join(" \xB7 "))),d.kind==="file"&&!b&&e.createElement(gt,null,i("unsupported")))})))}function De(){let{state:t,run:r,reset:o}=j(),n=H(),[s,a]=e.useState(null),l=e.useRef(null),f=e.useRef(!1),p=e.useRef(n),u=e.useCallback(c=>{a(c),r(c),window.setTimeout(()=>{l.current?.scrollIntoView({behavior:"smooth",block:"start"})},30)},[r]);e.useEffect(()=>{p.current!==n&&(p.current=n,s&&r(s))},[n,s,r]),e.useEffect(()=>{if(f.current)return;f.current=!0;let c=vt(),x=c.get("artifact"),m=c.get("path"),w=c.get("name")||void 0;x?u({filename:w||D(x),source:{kind:"artifact",url:x}}):m&&u({filename:w||D(m),source:{kind:"workspace",path:m,root:c.get("root")||"project"}})},[u]);let g=null;return t.status!=="idle"&&(g=e.createElement("section",{className:"qopx-card",ref:l},e.createElement(U,{state:t,onClose:o,onRetry:s?()=>void r(s):void 0,height:"70vh"}))),e.createElement("div",{className:"qopx-page","data-theme":n},e.createElement("header",{className:"qopx-hero"},e.createElement("h1",null,e.createElement("span",{className:"qopx-hero-icon","aria-hidden":"true"},e.createElement(fe,null)),i("routeTitle")),e.createElement("p",null,i("pageSubtitle"))),e.createElement("div",{className:"qopx-grid"},e.createElement("section",{className:"qopx-card"},e.createElement("div",{className:"qopx-card-body"},e.createElement(yt,{onFile:c=>u({filename:c.name,source:{kind:"upload",blob:c}})}),e.createElement(bt,null))),g,e.createElement("section",{className:"qopx-card"},e.createElement("div",{className:"qopx-card-head"},e.createElement("div",{className:"qopx-card-title"},e.createElement(me,null),i("workspace")),e.createElement(R.Text,{type:"secondary",style:{fontSize:12}},i("cacheNote"))),e.createElement(qt,{onPick:u}))))}var J="qwenpaw-office-preview";B();Te();try{window.QwenPaw.registerRoutes?.(J,[{path:`/plugin/${J}`,component:De,label:K("Office \u9884\u89C8","Office Preview"),title:K("Office \u6587\u6863\u9884\u89C8","Office Document Preview"),icon:"\u{1F4C4}",priority:46}])}catch(t){console.error(`[${J}] failed to register route`,t)}try{window.QwenPawOfficePreview={open:t=>{let r={source:t.artifact?"attachment":"workspace",path:t.path||"",artifactUrl:t.artifact,root:t.root||"project"};window.dispatchEvent(new CustomEvent("qwenpaw:open-file-preview",{detail:{target:r}}))},title:i("routeLabel")}}catch{}
