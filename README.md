# Office Viewer

使用 TypeScript 开发的 Pencode 只读 Office 预览插件，通过 peb v18 在工作区主机运行。作为独立 Git 子模块维护，主仓库保留插件版本引用。

- **Word `.docx`**：正文、标题、粗体、斜体、下划线、列表、表格和内嵌图片。
- **Excel `.xlsx`**：工作表切换、共享字符串、内联字符串、公式缓存值、常见日期与百分比格式；当前工作表连续滚动浏览全部行列，选中的工作表以主题强调色高亮。
- **PowerPoint `.pptx`**：按演示文稿顺序浏览幻灯片中的文本、表格和图片。

预览随编辑器主题变化，支持鼠标拖选文本、Ctrl+C 复制、Ctrl+A 全选和 Ctrl+F 搜索（macOS 使用 Cmd）；Excel 搜索覆盖当前工作表全部内容。磁盘文件变化时自动刷新，预览失败时可点击“重试”。注册为这三种文件的默认打开方式，整个过程不会修改 Office 文件，也不会上传文件或加载外部图片。

此版本提供内容预览，不保证 Office 的分页、字体、幻灯片坐标和复杂版式一致。未实现图表、SmartArt、公式计算、批注及修订追踪；Excel 使用已保存的公式缓存，缺少缓存时显示提示，隐藏工作表不展示，合并单元格按原始单元格显示。单元格保留完整文本供选择、复制与搜索。旧版 `.doc/.xls/.ppt`、宏格式和加密文档需要先转换为未加密的 OOXML。

限制：原文件 32 MiB、解压总量 64 MiB、4096 个 ZIP 项、单项 16 MiB、单个 XML 8 MiB。内嵌图片支持 PNG/JPEG/GIF/WebP，每张最多 256 KiB；无法预览的图片显示占位说明。单个预览视图约 1.8 MiB，超限时显示错误和重试按钮。

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

`npm run build && npm test` 覆盖三种文件解析、XML 转义、工作表与幻灯片切换、损坏文件重试、只读行为及实际 Node 子进程上的 peb 会话。主仓库的 `server/extensionhost` 测试还验证该插件可被 Go 宿主注册。

MIT；图标来自 Lucide，许可证见 `resources/LUCIDE-LICENSE.txt`。
