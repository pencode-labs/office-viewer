# Office Viewer

使用 TypeScript 开发的 Pencode 只读 Office 预览插件，通过 peb v20 在工作区主机运行。作为独立 Git 子模块维护，主仓库保留插件版本引用。

- **Word `.docx`**：用 docx-preview 按原稿纸张尺寸、页边距和分页标记展示页面，保留页眉页脚、段落样式、表格和内嵌图片。窄窗可横向滚动。超出布局预览容量时可分段阅读。
- **Excel `.xlsx`**：工作表切换、共享字符串、内联字符串、公式缓存值、常见日期与百分比格式；读取保存的列宽、行高、单元格对齐与自动换行设置，当前工作表建立轻量行索引，按可视行列解析并分块传输数据，以二维虚拟滚动浏览全部行列，只挂载可视区域及缓冲区，保留至少 100 行、26 列的空白网格范围，选中的工作表以主题强调色高亮。
- **PowerPoint `.pptx`**：用 pptx-glimpse 和 resvg 渲染幻灯片，按原比例居中展示。幻灯片及缩略图遇到缺失的中文字形时使用随插件分发的 Noto Sans SC。侧栏支持直接跳页，当前页附近按需渲染缩略图。点击侧栏底部的「文本视图」可选择、复制和搜索当前页文字。

预览随编辑器主题变化，支持鼠标拖选文本、Ctrl+C 复制、Ctrl+A 全选和 Ctrl+F 搜索（macOS 使用 Cmd）；Excel 查找、全选及复制时按需分块补齐当前工作表数据，搜索覆盖全部单元格，定位时按需挂载目标区域。Word 的搜索和全选覆盖当前预览内容，PowerPoint 的搜索和全选覆盖当前文本视图的内容。复制按单元格位置保留制表符和换行，全选包含离屏数据，不含行列标题。磁盘文件变化时自动刷新，预览失败时可点击“重试”。注册为这三种文件的默认打开方式，整个过程不会修改 Office 文件，也不会上传文件或加载外部图片。

Word 依赖文档保存的分页信息，自动分页可能与 Word 不同。PowerPoint 以静态图片展示，效果取决于本机安装的字体和渲染库支持的版式。Excel 使用已保存的公式缓存，缺少缓存时显示提示，隐藏工作表不展示，合并单元格按原始单元格显示。单元格保留完整文本供复制与搜索；虚拟网格使用保存的行高，换行文本超出该高度时裁剪，超长文本仅排版当前片段，搜索定位时切换到目标片段。旧版 `.doc/.xls/.ppt`、宏格式和加密文档需要先转换为未加密的 OOXML。

预览支持原文件最多 32 MiB、解压总量 64 MiB、4096 个 ZIP 项、单项 16 MiB、单个 XML 8 MiB。Word 支持最多 256 KiB 的 PNG 和 JPEG 内嵌图片，超限或不支持的图像显示占位说明。Excel 按需读取 ZIP 成员，每次传输最多 64 KiB，宿主完整表格索引的单次组装预算为 128 MiB。

## 构建和安装

要求 Node.js 22+（含 npm），或 Bun。源码依赖主仓库同版本的 `peb/ts`，尚未依赖公开 npm SDK。请在 Pencode 的 `extensions/office-viewer` 中开发；单独 clone 时也放到该位置。

```sh
npm ci
npm run build
npm test
```

使用 Bun：

```sh
bun install --frozen-lockfile
bun --bun run build
bun --bun run test
```

构建将 JavaScript 依赖打包到 `dist/office-viewer.mjs`，并将 PowerPoint 使用的 `resvg.wasm` 放在同一目录。Office 文件由工作区主机上的 JavaScript 运行时读取并渲染。

在 Pencode 主仓库执行 `make install-extensions` 自动构建安装，也可以在扩展管理器安装下述打包目录。分发目录仅需 `dist/`、`resources/`、`package.json` 和 `LICENSE`，可运行以下命令生成：

```sh
npm run package
```

将生成的 `package/` 目录交给扩展管理器安装，避免将开发依赖一并上传。远程项目的运行时及插件均安装在远端工作主机。服务端优先使用 PATH 中的 `node`，没有时使用 `bun`；入口无需可执行权限，各平台使用同一份打包产物。

## 开发验证

`npm run build && npm test` 覆盖三种文件解析、XML 转义、工作表与幻灯片切换、损坏文件重试、只读行为、超过原 1.8 MiB 限制的表格和文档、分块传输与实际 Node 子进程上的 peb 会话。主仓库的 `server/extensionhost` 测试还验证该插件可被 Go 宿主注册。

MIT；图标来自 Lucide，许可证见 `resources/LUCIDE-LICENSE.txt`。

## 自动发布与下载

推送 `v0.1.0` 等版本 tag 后，GitHub Actions 使用固定版本的 PEB v20 SDK，
安装锁定依赖、构建并测试，再自动创建 GitHub Release。预发布 tag（如
`v0.2.0-rc.1`）会创建预发布版本。

附件统一为 `office-viewer-<tag>-<os>-<arch>.zip`，支持 `linux`、`darwin`、
`windows`，各提供 `amd64`、`arm64`，并附 `SHA256SUMS`。各平台共用打包后的
JavaScript，ZIP 仅包含顶层 `office-viewer/` 目录，内有 `dist/`、`resources/`、
`package.json`、README 和许可证；清单及 package.json 版本与 tag 一致。

下载工作区服务端对应的平台包，直接解压到 `~/.pencode/extensions/`（Windows 为
`%USERPROFILE%/.pencode/extensions/`），得到
`~/.pencode/extensions/office-viewer/resources/extension.json`，重启 Pencode
即可发现扩展。远程工作区请在远端解压，也可在扩展管理器选择解压后的
`office-viewer/` 目录安装。运行主机需提供 Node.js 22+ 或 Bun，无需 npm、
node_modules、编译器或 Office。
