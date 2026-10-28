# QwenPaw Office 预览插件 · 离线安装包

版本：{{VERSION}}

这个压缩包已经内置：

- 构建好的前端 bundle（`dist/index.js`）；
- 纯 Python 依赖（`mammoth`、`openpyxl` 及其依赖）直接放在插件根目录；
- 后端转换器与预览页面。

安装和运行**不需要访问 PyPI、npm 或任何外部网络**。

## 安装

把 zip 拷到离线机器后执行：

```bash
qwenpaw plugin install ./qwenpaw-office-preview-offline-{{VERSION}}.zip
```

也可以先解压，再安装目录：

```bash
unzip qwenpaw-office-preview-offline-{{VERSION}}.zip
qwenpaw plugin install ./qwenpaw-office-preview
```

安装完成后重启 QwenPaw（或在运行中让 CLI 热加载），侧边栏
**设置 → 插件 → Office 预览** 即可使用。可以用下面的接口确认转换器状态：

```bash
curl -s http://127.0.0.1:8000/api/office-preview/health | jq
```

预期 `renderers.docx`、`renderers.xlsx` 为 `true`；`renderers.legacy_doc_ppt_xls`
取决于本机是否安装了 LibreOffice。

## 与联网安装版的差异

| 项目 | 联网安装版 | 离线安装包 |
| --- | --- | --- |
| Word（docx） | mammoth + python-docx 兜底 | mammoth |
| Excel（xlsx） | openpyxl | openpyxl |
| PowerPoint（pptx） | python-pptx | 内置纯 Python 解析器（`converter/pptx_native.py`） |
| 旧版 doc/ppt/xls | LibreOffice | LibreOffice（仍需单独安装） |
| 依赖安装 | 安装时访问 PyPI | 无需安装，已内置 |

内置 PPT 解析器支持文本、占位符、表格、图片、分组形状、简单图表和演讲者备注；
SmartArt、动画、音视频和 OLE 对象不会渲染。需要更高保真度时，可以在联网机器上
使用源码目录安装，让 QwenPaw 安装 `python-pptx`。

## 卸载

```bash
qwenpaw plugin uninstall qwenpaw-office-preview
```
