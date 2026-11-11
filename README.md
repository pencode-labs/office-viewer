# Office Viewer

使用 TypeScript 开发的 Pencode 只读 Office 预览插件，通过 peb v18 在工作区主机运行。作为独立 Git 子模块维护，主仓库保留插件版本引用。

- **Word `.docx`**：正文、标题、粗体、斜体、下划线、列表、表格和内嵌图片；正文按需解析、分段渲染，大段落和大表格也能分段浏览。
- **Excel `.xlsx`**：工作表切换、共享字符串、内联字符串、公式缓存值、常见日期与百分比格式；读取保存的列宽、行高、单元格对齐与自动换行设置，当前工作表建立轻量行索引，按可视行列解析并分块传输数据，以二维虚拟滚动浏览全部行列，只挂载可视区域及缓冲区，保留至少 100 行、26 列的空白网格范围，选中的工作表以主题强调色高亮。
- **PowerPoint `.pptx`**：按演示文稿顺序按需加载幻灯片中的文本、表格和图片，较长幻灯片可分段浏览。

预览随编辑器主题变化，支持鼠标拖选文本、Ctrl+C 复制、Ctrl+A 全选和 Ctrl+F 搜索（macOS 使用 Cmd）；Excel 查找、全选及复制时按需分块补齐当前工作表数据，搜索覆盖全部单元格，定位时按需挂载目标区域。Word 和 PowerPoint 的搜索、全选作用于当前内容分段。复制按单元格位置保留制表符和换行，全选包含离屏数据，不含行列标题。磁盘文件变化时自动刷新，预览失败时可点击“重试”。注册为这三种文件的默认打开方式，整个过程不会修改 Office 文件，也不会上传文件或加载外部图片。

此版本提供内容预览，不保证 Office 的分页、字体、幻灯片坐标和复杂版式一致。未实现图表、SmartArt、公式计算、批注及修订追踪；Excel 使用已保存的公式缓存，缺少缓存时显示提示，隐藏工作表不展示，合并单元格按原始单元格显示。单元格保留完整文本供复制与搜索；虚拟网格使用保存的行高，换行文本超出该高度时裁剪，超长文本仅排版当前片段，搜索定位时切换到目标片段。旧版 `.doc/.xls/.ppt`、宏格式和加密文档需要先转换为未加密的 OOXML。

限制：原文件 32 MiB、解压总量 64 MiB、4096 个 ZIP 项、单项 16 MiB、单个 XML 8 MiB。内嵌图片支持 PNG/JPEG/GIF/WebP，每张最多 256 KiB；无法预览的图片显示占位说明。已移除约 1.8 MiB 的整页预览限制。ZIP 仅索引目录，按需读取和解压所用成员；工作表和共享字符串以源范围索引保存，单元格按需解析。Excel 每次传输最多 64 KiB，Word/PowerPoint 保留有限的已渲染分段缓存；PEB 消息大小限制不变，宿主完整表格索引的单次组装预算为 128 MiB。此功能需要同时更新 Pencode 主程序和插件。

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

构建生成 `dist/office-viewer.mjs`，包含 SDK、XML 和 ZIP 依赖。运行时仅需 JavaScript 运行时；不需要 `node_modules`、Office、LibreOffice 或浏览器。

在 Pencode 主仓库执行 `make install-extensions` 自动构建安装，也可以在扩展管理器安装下述打包目录。分发目录仅需 `dist/`、`resources/`、`package.json` 和 `LICENSE`，可运行以下命令生成：

```sh
npm run package
```

将生成的 `package/` 目录交给扩展管理器安装，避免将开发依赖一并上传。远程项目的运行时及插件均安装在远端工作主机。服务端优先使用 PATH 中的 `node`，没有时使用 `bun`；入口无需可执行权限，各平台使用同一份打包产物。

## 开发验证

`npm run build && npm test` 覆盖三种文件解析、XML 转义、工作表与幻灯片切换、损坏文件重试、只读行为、超过原 1.8 MiB 限制的表格和文档、分块传输与实际 Node 子进程上的 peb 会话。主仓库的 `server/extensionhost` 测试还验证该插件可被 Go 宿主注册。

MIT；图标来自 Lucide，许可证见 `resources/LUCIDE-LICENSE.txt`。
