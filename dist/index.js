var x=window.QwenPaw.host,e=x.React,Z=x.ReactDOM,h=x.antd,L=x.antdIcons;function b(){try{if(document.documentElement.classList.contains("dark-mode"))return"dark";x.useTheme}catch{}return"light"}function ye(){try{let t=window.sessionStorage.getItem("qwenpaw-agent-storage")??window.localStorage.getItem("qwenpaw-agent-storage");if(t){let o=JSON.parse(t)?.state?.selectedAgent;if(typeof o=="string"&&o)return o}}catch{}try{return x.getSelectedAgentId?.()??"default"}catch{return"default"}}function X(){let t={};try{let o=x.getApiToken();o&&(t.Authorization=`Bearer ${o}`)}catch{}let r=ye();return r&&(t["X-Agent-Id"]=r),t}async function E(t,r={}){if(x.fetch)return x.fetch(t,r);let o=/^https?:\/\//i.test(t)?t:x.getApiUrl(t);return fetch(o,{...r,headers:{...X(),...r.headers??{}}})}async function Y(t,r={}){return fetch(t,{...r,headers:{...X(),...r.headers??{}}})}var J=new Set(["csv","doc","docm","docx","dot","dotm","dotx","dps","et","pps","ppsx","ppt","pptm","pptx","potx","tsv","wps","xls","xlsm","xlsx","xltm","xltx"]),ee=Array.from(J).map(t=>`.${t}`).join(","),k=class extends Error{constructor(r,o,n){super(r),this.name="ApiError",this.status=o,this.code=n}};function be(t){let o=(t||"").split(/[?#]/,1)[0].replace(/\\/g,"/").split("/").pop()||"",n=o.lastIndexOf(".");return n>=0?o.slice(n+1).toLowerCase():""}function q(t){return J.has(be(t))}function P(t){return(t||"").replace(/\\/g,"/").split(/[?#]/,1)[0].split("/").filter(Boolean).pop()||"document"}async function D(t){let r=t.status,o=`${r} ${t.statusText}`.trim(),n=`http_${r}`;try{let l=await t.json(),a=l?.detail??l;typeof a=="string"?o=a:a&&typeof a=="object"&&(typeof a.message=="string"&&(o=a.message),typeof a.code=="string"&&(n=a.code))}catch{try{let l=await t.text();l&&(o=l.slice(0,500))}catch{}}return new k(o,r,n)}async function te(){let t=await E("/office-preview/health");if(!t.ok)throw await D(t);return await t.json()}async function re(t,r,o=400){let n=new URLSearchParams({path:t,root:r,limit:String(o)}),l=await E(`/workspace/tree?${n.toString()}`);if(!l.ok)throw await D(l);return await l.json()}async function oe(t){let{source:r}=t;if(r.kind==="upload")return{blob:r.blob,filename:t.filename};let o;if(r.kind==="workspace"){let a=new URLSearchParams({path:r.path,root:r.root||"project"});o=await E(`/workspace/file-download?${a.toString()}`)}else o=await Y(r.url);if(!o.ok)throw await D(o);let n=await o.blob(),l=t.filename||P(r.kind==="workspace"?r.path:r.path||"");return{blob:n,filename:l}}async function ne(t,r,o){let n=new URLSearchParams({filename:r||"document.docx",theme:b()}),l=await E(`/office-preview/convert?${n.toString()}`,{method:"POST",headers:{"Content-Type":"application/octet-stream"},body:t,signal:o});if(!l.ok)throw await D(l);let a=await l.json();if(!a||typeof a.html!="string"||!a.html)throw new k("The server returned an empty preview.",500,"empty_result");return a}var ke={routeLabel:"Office \u9884\u89C8",routeTitle:"Office \u6587\u6863\u9884\u89C8",pageSubtitle:"\u65E0\u9700\u4E0B\u8F7D\u5373\u53EF\u67E5\u770B Word\u3001PowerPoint \u548C Excel \u6587\u6863\u3002\u62D6\u5165\u672C\u5730\u6587\u4EF6\uFF0C\u6216\u76F4\u63A5\u4ECE\u5F53\u524D\u5DE5\u4F5C\u533A\u6253\u5F00\u3002",dropTitle:"\u62D6\u62FD\u6587\u4EF6\u5230\u8FD9\u91CC\uFF0C\u6216\u70B9\u51FB\u9009\u62E9",dropHint:"\u652F\u6301 doc / docx / ppt / pptx / xls / xlsx\uFF0C\u4EE5\u53CA csv / tsv",chooseFile:"\u9009\u62E9\u6587\u4EF6",workspace:"\u5DE5\u4F5C\u533A\u6587\u4EF6",projectRoot:"\u9879\u76EE\u76EE\u5F55",workspaceRoot:"\u5DE5\u4F5C\u533A\u76EE\u5F55",refresh:"\u5237\u65B0",filterOnly:"\u4EC5\u663E\u793A\u53EF\u9884\u89C8",emptyDir:"\u8FD9\u4E2A\u76EE\u5F55\u91CC\u6CA1\u6709\u53EF\u9884\u89C8\u7684\u6587\u6863\u3002",loadDirFailed:"\u8BFB\u53D6\u76EE\u5F55\u5931\u8D25",unsupported:"\u4E0D\u652F\u6301\u9884\u89C8\u8BE5\u6587\u4EF6\u7C7B\u578B",unsupportedHint:"\u652F\u6301 Word\u3001PowerPoint\u3001Excel \u6587\u6863\u4EE5\u53CA CSV\u3002",loading:"\u6B63\u5728\u751F\u6210\u9884\u89C8\u2026",converting:"\u6B63\u5728\u89E3\u6790\u6587\u6863\uFF0C\u8F83\u5927\u7684\u6587\u4EF6\u53EF\u80FD\u9700\u8981\u51E0\u79D2\u949F\u3002",retry:"\u91CD\u8BD5",close:"\u5173\u95ED",download:"\u4E0B\u8F7D\u539F\u6587\u4EF6",openWindow:"\u65B0\u7A97\u53E3\u6253\u5F00",previewFailed:"\u9884\u89C8\u5931\u8D25",backToList:"\u8FD4\u56DE\u6587\u4EF6\u5217\u8868",size:"\u5927\u5C0F",modified:"\u4FEE\u6539\u65F6\u95F4",status:"\u8F6C\u6362\u5668\u72B6\u6001",statusOk:"\u53EF\u7528",statusMissing:"\u672A\u5B89\u88C5",healthFailed:"\u65E0\u6CD5\u83B7\u53D6\u8F6C\u6362\u5668\u72B6\u6001",depsMissing:"\u90E8\u5206\u8F6C\u6362\u5668\u4F9D\u8D56\u7F3A\u5931\uFF0C\u8BF7\u91CD\u65B0\u5B89\u88C5\u63D2\u4EF6\u4EE5\u81EA\u52A8\u8865\u9F50\u4F9D\u8D56\u3002",legacyHint:"\u65E7\u7248 .doc / .ppt / .xls \u6587\u4EF6\u9700\u8981\u672C\u673A\u5B89\u88C5 LibreOffice\uFF1B\u4E5F\u53EF\u4EE5\u5148\u53E6\u5B58\u4E3A docx / pptx / xlsx\u3002",cacheNote:"\u8F6C\u6362\u7ED3\u679C\u4F1A\u7F13\u5B58\u5728\u5185\u5B58\u4E2D\uFF0C\u91CD\u590D\u6253\u5F00\u540C\u4E00\u6587\u4EF6\u4F1A\u66F4\u5FEB\u3002",fileTooLarge:"\u6587\u4EF6\u8FC7\u5927\uFF0C\u65E0\u6CD5\u5728\u7EBF\u9884\u89C8\u3002",dialogLabel:"Office \u6587\u6863\u9884\u89C8",navigate:"\u5B9A\u4F4D",parent:"\u4E0A\u4E00\u7EA7",root:"\u6839\u76EE\u5F55",entries:"\u9879",openInPage:"\u5728\u72EC\u7ACB\u9875\u9762\u4E2D\u6253\u5F00",emptyState:"\u9009\u62E9\u5DE6\u4FA7\u6587\u4EF6\u6216\u62D6\u5165\u6587\u6863\u5F00\u59CB\u9884\u89C8",convertedFrom:"\u5DF2\u4ECE\u65E7\u683C\u5F0F\u8F6C\u6362",warnings:"\u683C\u5F0F\u63D0\u793A",timeout:"\u9884\u89C8\u5DF2\u53D6\u6D88\u6216\u8D85\u65F6\u3002"},qe={routeLabel:"Office Preview",routeTitle:"Office Document Preview",pageSubtitle:"Read Word, PowerPoint and Excel documents without downloading them. Drop a local file or open one from the workspace.",dropTitle:"Drop a file here, or click to choose",dropHint:"Supports doc / docx / ppt / pptx / xls / xlsx, plus csv / tsv",chooseFile:"Choose file",workspace:"Workspace files",projectRoot:"Project",workspaceRoot:"Workspace",refresh:"Refresh",filterOnly:"Previewable only",emptyDir:"No previewable documents in this folder.",loadDirFailed:"Failed to read the folder",unsupported:"This file type cannot be previewed",unsupportedHint:"Word, PowerPoint, Excel and CSV documents are supported.",loading:"Building preview\u2026",converting:"Parsing the document; large files may take a few seconds.",retry:"Retry",close:"Close",download:"Download original",openWindow:"Open in new window",previewFailed:"Preview failed",backToList:"Back to files",size:"Size",modified:"Modified",status:"Converter status",statusOk:"Available",statusMissing:"Not installed",healthFailed:"Could not read converter status",depsMissing:"Some converters are unavailable. Reinstall the plugin to restore its Python dependencies.",legacyHint:"Legacy .doc / .ppt / .xls previews require LibreOffice on the host; alternatively save the file as docx / pptx / xlsx.",cacheNote:"Converted documents are cached in memory, so reopening the same file is faster.",fileTooLarge:"The file is too large to preview online.",dialogLabel:"Office document preview",navigate:"Go",parent:"Up",root:"Root",entries:"items",openInPage:"Open in dedicated page",emptyState:"Pick a file on the left or drop a document to begin",convertedFrom:"Converted from a legacy format",warnings:"Document notes",timeout:"Preview was cancelled or timed out."};function Pe(){try{return(window.localStorage.getItem("language")||document.documentElement.getAttribute("lang")||navigator.language||"en").toLowerCase().startsWith("zh")?"zh":"en"}catch{return"zh"}}function i(t){return(Pe()==="zh"?ke:qe)[t]}function $(t,r){return o=>String(o||"").toLowerCase().startsWith("zh")?t:r}function z(t){if(t==null||!Number.isFinite(t))return"";let r=["B","KB","MB","GB"],o=t,n=0;for(;o>=1024&&n<r.length-1;)o/=1024,n+=1;let l=o>=100||n===0?0:1;return`${o.toFixed(l)} ${r[n]}`}function ie(t){if(!t)return"";let r=new Date(t);if(Number.isNaN(r.getTime()))return"";try{return r.toLocaleString(void 0,{year:"numeric",month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit"})}catch{return r.toISOString()}}function I(t){let r=(t.split(".").pop()||"").toLowerCase();return["doc","docx","docm","dot","dotx","dotm","wps"].includes(r)?"\u{1F4DD}":["ppt","pptx","pptm","pps","ppsx","potx","dps"].includes(r)?"\u{1F4CA}":["xls","xlsx","xlsm","xltx","xltm","et","csv","tsv"].includes(r)?"\u{1F4C8}":"\u{1F4C4}"}var{Alert:ae,Button:S,Space:Te,Spin:Ee,Tag:W,Tooltip:N,Typography:F}=h,{CloseOutlined:Se,DownloadOutlined:Ne,ExportOutlined:Oe,FileWarningOutlined:Ce,ReloadOutlined:se}=L;function Le(t,r){if(!t)return;let o=URL.createObjectURL(t),n=document.createElement("a");n.href=o,n.download=r||"document",document.body.appendChild(n),n.click(),n.remove(),window.setTimeout(()=>URL.revokeObjectURL(o),1e4)}function De(t){let r=new Blob([t],{type:"text/html;charset=utf-8"}),o=URL.createObjectURL(r);if(!window.open(o,"_blank","noopener,noreferrer")){URL.revokeObjectURL(o);return}window.setTimeout(()=>URL.revokeObjectURL(o),6e4)}function B({state:t,onClose:r,onRetry:o,height:n="72vh",className:l}){let{status:a,result:s,filename:p,error:f}=t,d=s?.meta?.label||s?.kind?.toUpperCase()||"",w=s?.warnings??[],u=s?.meta?.converted_from,m=e.createElement("div",{className:"qopx-viewer-head"},e.createElement("div",{className:"qopx-viewer-title"},e.createElement("span",{"aria-hidden":"true"},I(p)),e.createElement("span",{title:p},p||i("routeTitle")),a==="ready"&&d&&e.createElement(W,{color:"blue"},d),u&&e.createElement(N,{title:i("convertedFrom")},e.createElement(W,null,String(u).replace(".","").toUpperCase())),s?.cached&&e.createElement(W,{color:"default"},"cache"),s?.meta?.size?e.createElement(F.Text,{type:"secondary",style:{fontSize:12}},z(s.meta.size)):null),e.createElement(Te,{size:6,wrap:!0},a==="ready"&&s&&e.createElement(e.Fragment,null,t.blob&&e.createElement(N,{title:i("download")},e.createElement(S,{size:"small","aria-label":i("download"),icon:e.createElement(Ne,null),onClick:()=>Le(t.blob,p)})),e.createElement(N,{title:i("openWindow")},e.createElement(S,{size:"small","aria-label":i("openWindow"),icon:e.createElement(Oe,null),onClick:()=>De(s.html)}))),o&&a!=="loading"&&e.createElement(N,{title:i("retry")},e.createElement(S,{size:"small","aria-label":i("retry"),icon:e.createElement(se,null),onClick:o})),r&&e.createElement(N,{title:i("close")},e.createElement(S,{size:"small",type:"text","aria-label":i("close"),icon:e.createElement(Se,null),onClick:r})))),g;return a==="loading"?g=e.createElement("div",{className:"qopx-viewer-state"},e.createElement(Ee,{size:"large"}),e.createElement(F.Text,{strong:!0},i("loading")),e.createElement(F.Text,{type:"secondary"},i("converting"))):a==="error"?g=e.createElement("div",{style:{padding:22}},e.createElement(ae,{type:"error",showIcon:!0,icon:e.createElement(Ce,null),message:i("previewFailed"),description:e.createElement("div",null,e.createElement("p",{style:{margin:"0 0 10px"}},f),t.errorCode==="legacy_converter_missing"&&e.createElement("p",{style:{margin:0,opacity:.75}},i("legacyHint")),t.errorCode==="missing_dependency"&&e.createElement("p",{style:{margin:0,opacity:.75}},i("depsMissing")),o&&e.createElement(S,{size:"small",style:{marginTop:12},icon:e.createElement(se,null),onClick:o},i("retry")))})):a==="ready"&&s?g=e.createElement(e.Fragment,null,w.length>0&&e.createElement(ae,{type:"warning",showIcon:!0,style:{borderRadius:0,borderInline:0,borderTop:0},message:i("warnings"),description:e.createElement("ul",{className:"qopx-warning-list"},w.slice(0,6).map((y,T)=>e.createElement("li",{key:`${T}-${y.slice(0,24)}`},y)))}),e.createElement("iframe",{className:"qopx-viewer-frame",title:p||i("routeTitle"),srcDoc:s.html,sandbox:"",referrerPolicy:"no-referrer"})):g=e.createElement("div",{className:"qopx-viewer-state"},e.createElement(F.Text,{type:"secondary"},i("emptyState"))),e.createElement("div",{className:`qopx-viewer${l?` ${l}`:""}`,style:{height:n}},m,e.createElement("div",{className:"qopx-viewer-body"},g))}var le="qwenpaw-office-preview-styles",ze=`
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

.qopx-overlay {
  position: fixed; inset: 0; z-index: 2147483000;
  display: flex; align-items: center; justify-content: center;
  padding: 22px;
  background: rgba(15, 23, 42, .46);
  backdrop-filter: blur(3px);
}
.qopx-overlay-panel {
  display: flex; flex-direction: column;
  width: min(1420px, 96vw); height: min(92vh, 1000px);
  background: var(--ant-color-bg-container, #fff);
  color: var(--ant-color-text, inherit);
  border-radius: 14px; overflow: hidden;
  box-shadow: 0 30px 80px rgba(15, 23, 42, .32);
  border: 1px solid var(--ant-color-border-secondary, #e5e7eb);
}
.qopx-overlay-panel .qopx-viewer { flex: 1; }
.qopx-viewer-body { flex: 1; min-height: 0; display: flex; flex-direction: column; }
.qopx-warning-list { margin: 0; padding-left: 18px; }

@media (max-width: 720px) {
  .qopx-page { padding: 14px 12px 28px; }
  .qopx-overlay { padding: 0; }
  .qopx-overlay-panel { width: 100vw; height: 100vh; border-radius: 0; }
  .qopx-file-list { max-height: 52vh; }
}
`;function j(){if(typeof document>"u"||document.getElementById(le))return;let t=document.createElement("style");t.id=le,t.textContent=ze,document.head.appendChild(t)}var{useCallback:R,useEffect:Ie,useRef:pe,useState:Fe}=e,ce={status:"idle",filename:""},Be=24e4;function A(){let[t,r]=Fe(ce),o=pe(0),n=pe(null),l=R(()=>{o.current+=1,n.current?.abort(),n.current=null},[]),a=R(async p=>{let f=o.current+1;o.current=f,n.current?.abort();let d=new AbortController;n.current=d,r({status:"loading",filename:p.filename});let w=window.setTimeout(()=>d.abort(),Be);try{let u=await oe(p);if(f!==o.current)return;let m=u.filename||p.filename;r({status:"loading",filename:m});let g=await ne(u.blob,m,d.signal);if(f!==o.current)return;r({status:"ready",filename:m,result:g,blob:u.blob})}catch(u){if(f!==o.current)return;if(u?.name==="AbortError"){r({status:"error",filename:p.filename,error:i("timeout"),errorCode:"aborted"});return}let m=u instanceof k?u:null,g=m?.status===413?i("fileTooLarge"):u?.message||i("previewFailed");r({status:"error",filename:p.filename,error:g,errorCode:m?.code})}finally{window.clearTimeout(w),n.current===d&&(n.current=null)}},[]),s=R(()=>{l(),r(ce)},[l]);return Ie(()=>()=>l(),[l]),{state:t,run:a,reset:s}}var{ConfigProvider:de}=h;function je({request:t,onClose:r}){let{state:o,run:n}=A();e.useEffect(()=>{n(t)},[]),e.useEffect(()=>{let p=f=>{f.key==="Escape"&&(f.stopPropagation(),r())};return window.addEventListener("keydown",p,!0),()=>window.removeEventListener("keydown",p,!0)},[r]);let a=b()==="dark"&&h.theme?.darkAlgorithm?h.theme.darkAlgorithm:h.theme?.defaultAlgorithm,s=e.createElement("div",{className:"qopx-overlay",onMouseDown:p=>{p.target===p.currentTarget&&r()}},e.createElement("div",{className:"qopx-overlay-panel",role:"dialog","aria-modal":"true","aria-label":i("dialogLabel")},e.createElement(B,{state:o,onClose:r,onRetry:()=>void n(t),height:"100%"})));return de?e.createElement(de,{theme:a?{algorithm:a}:void 0},s):s}var O=null,H=null,ue=0,K="",M=!1;function Ae(){return O||(O=document.createElement("div"),O.id="qwenpaw-office-preview-overlay",document.body.appendChild(O),H=Z.createRoot(O)),H}function fe(t){j();let r=Ae();ue+=1,M||(K=document.body.style.overflow,M=!0),document.body.style.overflow="hidden",r.render(e.createElement(je,{key:`preview-${ue}`,request:t,onClose:He}))}function He(){M&&(M=!1,H&&H.render(null),document.body.style.overflow=K,K="")}var Me="qwenpaw:open-file-preview";function Ue(t){return t.replace(/^\/+/,"").split("/").map(r=>encodeURIComponent(r)).join("/")}function _e(t){let r=(t.path||"").trim(),o=(t.artifactUrl||"").trim();if(!r&&!o||!q(r||o))return null;let n=P(r||o);return t.source==="workspace"&&r?{filename:n,source:{kind:"workspace",path:r,root:t.root||"project"}}:o?{filename:n,source:{kind:"artifact",url:o,path:r}}:r.startsWith("/")||/^[a-zA-Z]:[\\/]/.test(r)?{filename:n,source:{kind:"artifact",url:x.getApiUrl(`/files/preview${Ue(r)}`),path:r}}:r?{filename:n,source:{kind:"workspace",path:r,root:t.root||"project"}}:null}function $e(t){let o=t.detail?.target;if(!o)return;let n=_e(o);n&&(t.preventDefault(),t.stopImmediatePropagation(),fe(n))}var ge=!1;function me(){ge||(ge=!0,window.addEventListener(Me,$e,!0))}var{Alert:we,Button:V,Empty:xe,Radio:We,Space:Re,Spin:Ke,Switch:Ve,Tag:Ge,Typography:U}=h,{ReloadOutlined:Qe,UploadOutlined:Ze}=L;function Xe(){let t=window.location.hash||"",r=t.indexOf("?");return r>=0?new URLSearchParams(t.slice(r+1)):new URLSearchParams(window.location.search)}function Ye(){let[t,r]=e.useState(null),[o,n]=e.useState(!1);if(e.useEffect(()=>{let s=!0;return te().then(p=>{s&&r(p)}).catch(()=>{s&&n(!0)}),()=>{s=!1}},[]),o)return e.createElement(U.Text,{type:"secondary",style:{fontSize:12}},i("healthFailed"));if(!t)return null;let l=[{label:"Word",ok:!!t.renderers.docx},{label:"PowerPoint",ok:!!t.renderers.pptx},{label:"Excel",ok:!!t.renderers.xlsx},{label:"LibreOffice",ok:!!t.libreoffice}],a=l.some(s=>!s.ok&&s.label!=="LibreOffice");return e.createElement(e.Fragment,null,e.createElement("div",{className:"qopx-status"},e.createElement("span",{className:"qopx-chip"},i("status")),l.map(s=>e.createElement("span",{className:"qopx-chip",key:s.label},e.createElement("span",{className:`qopx-dot ${s.ok?"":"qopx-bad"}`,"aria-hidden":"true"}),s.label,": ",s.ok?i("statusOk"):i("statusMissing")))),a&&e.createElement(we,{type:"warning",showIcon:!0,style:{marginTop:12},message:i("depsMissing"),description:t.notes?.length?e.createElement("ul",{className:"qopx-warning-list"},t.notes.map(s=>e.createElement("li",{key:s},s))):void 0}),!t.libreoffice&&e.createElement(U.Paragraph,{type:"secondary",style:{fontSize:12,margin:"10px 0 0"}},i("legacyHint")))}function Je({onFile:t}){let r=e.useRef(null),[o,n]=e.useState(!1),l=a=>{let s=a&&a[0];if(s){if(!q(s.name)){h.message.warning(`${i("unsupported")} \xB7 ${i("unsupportedHint")}`);return}t(s)}};return e.createElement("div",{className:`qopx-drop${o?" qopx-dragging":""}`,role:"button",tabIndex:0,onClick:()=>r.current?.click(),onKeyDown:a=>{(a.key==="Enter"||a.key===" ")&&(a.preventDefault(),r.current?.click())},onDragOver:a=>{a.preventDefault(),n(!0)},onDragLeave:a=>{a.preventDefault(),n(!1)},onDrop:a=>{a.preventDefault(),n(!1),l(a.dataTransfer?.files??null)}},e.createElement("span",{className:"qopx-drop-icon","aria-hidden":"true"},"\u{1F4C4}"),e.createElement("span",{className:"qopx-drop-title"},i("dropTitle")),e.createElement("span",{className:"qopx-drop-hint"},i("dropHint")),e.createElement("div",{className:"qopx-drop-actions"},e.createElement(V,{type:"primary",size:"small",icon:e.createElement(Ze,null),onClick:a=>{a.stopPropagation(),r.current?.click()}},i("chooseFile"))),e.createElement("input",{ref:r,type:"file",accept:ee,style:{display:"none"},onChange:a=>{l(a.target.files),a.target.value=""}}))}function et({onPick:t}){let[r,o]=e.useState("project"),[n,l]=e.useState(""),[a,s]=e.useState(null),[p,f]=e.useState(!1),[d,w]=e.useState(""),[u,m]=e.useState(!0),g=e.useRef(0),y=e.useCallback(async(c,v)=>{let C=g.current+1;g.current=C,f(!0),w("");try{let _=await re(c,v);C===g.current&&s(_)}catch(_){C===g.current&&(s(null),w(_?.message||i("loadDirFailed")))}finally{C===g.current&&f(!1)}},[]);e.useEffect(()=>{y(n,r)},[n,r,y]);let T=n.split("/").filter(Boolean),Q=(a?.entries??[]).filter(c=>c.kind==="directory"?!0:!u||q(c.path)),ve=()=>l(T.slice(0,-1).join("/"));return e.createElement(e.Fragment,null,e.createElement("div",{className:"qopx-breadcrumb"},e.createElement(We.Group,{size:"small",value:r,optionType:"button",options:[{label:i("projectRoot"),value:"project"},{label:i("workspaceRoot"),value:"workspace"}],onChange:c=>{l(""),o(c.target.value)}}),e.createElement("span",{style:{width:8}}),e.createElement("button",{type:"button",className:"qopx-crumb",onClick:()=>l(""),title:i("root")},r==="workspace"?i("workspaceRoot"):i("projectRoot")),T.map((c,v)=>e.createElement(e.Fragment,{key:`${c}-${v}`},e.createElement("span",{className:"qopx-crumb-sep"},"/"),e.createElement("button",{type:"button",className:"qopx-crumb",onClick:()=>l(T.slice(0,v+1).join("/"))},c))),e.createElement("span",{style:{flex:1}}),e.createElement(Re,{size:6},n&&e.createElement(V,{size:"small",type:"text",onClick:ve},i("parent")),e.createElement(V,{size:"small",type:"text",icon:e.createElement(Qe,null),onClick:()=>void y(n,r)},i("refresh")),e.createElement(U.Text,{type:"secondary",style:{fontSize:12}},i("filterOnly")),e.createElement(Ve,{size:"small",checked:u,onChange:m}))),d&&e.createElement("div",{style:{padding:16}},e.createElement(we,{type:"error",showIcon:!0,message:i("loadDirFailed"),description:d})),e.createElement("div",{className:"qopx-file-list"},p?e.createElement("div",{className:"qopx-empty"},e.createElement(Ke,{size:"small"})," ",e.createElement("span",{style:{marginLeft:8}},i("loading"))):Q.length===0?e.createElement("div",{className:"qopx-empty"},e.createElement(xe,{image:xe.PRESENTED_IMAGE_SIMPLE,description:i("emptyDir")})):Q.map(c=>{let v=c.kind==="directory"||q(c.path);return e.createElement("button",{type:"button",key:c.path,className:"qopx-file-row",disabled:!v,onClick:()=>{if(c.kind==="directory"){l(c.path);return}t({filename:c.name,source:{kind:"workspace",path:c.path,root:r}})}},e.createElement("span",{className:"qopx-file-icon","aria-hidden":"true"},c.kind==="directory"?"\u{1F4C1}":I(c.name)),e.createElement("span",{className:"qopx-file-main"},e.createElement("span",{className:"qopx-file-name",title:c.name},c.name),e.createElement("span",{className:"qopx-file-sub"},c.kind==="directory"?"folder":[z(c.size),ie(c.modified_at)].filter(Boolean).join(" \xB7 "))),c.kind==="file"&&!v&&e.createElement(Ge,null,i("unsupported")))})))}function he(){let{state:t,run:r,reset:o}=A(),[n,l]=e.useState(null),a=e.useRef(null),s=e.useRef(!1),p=e.useCallback(d=>{l(d),r(d),window.setTimeout(()=>{a.current?.scrollIntoView({behavior:"smooth",block:"start"})},30)},[r]);e.useEffect(()=>{if(s.current)return;s.current=!0;let d=Xe(),w=d.get("artifact"),u=d.get("path"),m=d.get("name")||void 0;w?p({filename:m||P(w),source:{kind:"artifact",url:w}}):u&&p({filename:m||P(u),source:{kind:"workspace",path:u,root:d.get("root")||"project"}})},[p]);let f=null;return t.status!=="idle"&&(f=e.createElement("section",{className:"qopx-card",ref:a},e.createElement(B,{state:t,onClose:o,onRetry:n?()=>void r(n):void 0,height:"70vh"}))),e.createElement("div",{className:"qopx-page","data-theme":b()},e.createElement("header",{className:"qopx-hero"},e.createElement("h1",null,e.createElement("span",{className:"qopx-hero-icon","aria-hidden":"true"},"\u{1F4C4}"),i("routeTitle")),e.createElement("p",null,i("pageSubtitle"))),e.createElement("div",{className:"qopx-grid"},e.createElement("section",{className:"qopx-card"},e.createElement("div",{className:"qopx-card-body"},e.createElement(Je,{onFile:d=>p({filename:d.name,source:{kind:"upload",blob:d}})}),e.createElement(Ye,null))),f,e.createElement("section",{className:"qopx-card"},e.createElement("div",{className:"qopx-card-head"},e.createElement("div",{className:"qopx-card-title"},"\u{1F5C2} ",i("workspace")),e.createElement(U.Text,{type:"secondary",style:{fontSize:12}},i("cacheNote"))),e.createElement(et,{onPick:p}))))}var G="qwenpaw-office-preview";j();me();try{window.QwenPaw.registerRoutes?.(G,[{path:`/plugin/${G}`,component:he,label:$("Office \u9884\u89C8","Office Preview"),title:$("Office \u6587\u6863\u9884\u89C8","Office Document Preview"),icon:"\u{1F4C4}",priority:46}])}catch(t){console.error(`[${G}] failed to register route`,t)}try{window.QwenPawOfficePreview={open:t=>{let r={source:t.artifact?"attachment":"workspace",path:t.path||"",artifactUrl:t.artifact,root:t.root||"project"};window.dispatchEvent(new CustomEvent("qwenpaw:open-file-preview",{detail:{target:r}}))},title:i("routeLabel")}}catch{}
