本项目已被dshfind.com收录

# dsh-better-sidebar-plugin-office

[![npm version](https://img.shields.io/npm/v/@huanlin/dsh-plugin-better-sidebar-plugin-office)](https://www.npmjs.com/package/@huanlin/dsh-plugin-better-sidebar-plugin-office)

DSH web 插件：为 better-sidebar 的编辑器提供 Office 三件套文件预览（`.docx` / `.xlsx` / `.pptx`）。

## 背景

better-sidebar 曾内置这三个 viewer，但 docx-preview / Univer（+ SheetJS）/ pptx-renderer 会把 client bundle 撑到约 23MB。本插件把它们拆成独立 bundle，better-sidebar 本体不再内联 Office 渲染库（约 5.9MB 全套 → 2.2MB 压缩），需要 Office 预览的用户单独安装本插件即可。

## 功能

- 通过 `ctx.betterSidebar.registerFileViewer` 注册 3 个 viewer（id 与内置一致）：
  - `docx`：`.docx`，docx-preview 保真渲染（样式/图片/表格，Alt+滚轮缩放）
  - `xlsx`：`.xlsx`，Univer sheets 预设（数据 + 公式 + 合并单元格 + 列宽/行高）
  - `pptx`：`.pptx`，@aiden0z/pptx-renderer 浏览器原生预览（翻页导航）
- viewer 描述符形状与原内置完全一致（`mediaUrl` 策略、priority 0、title/icon），因此：
  - 既有文件路由行为不变（扩展名匹配优先于 binary-download / code 兜底）
  - Side card 设置页自动显示这三个 viewer 的启用开关
  - 加载/渲染失败回退到「下载查看」链接
- 客户端组件自包含：自带 locale（中/英）与 CSS，不依赖 better-sidebar 的 client 内部实现

## 依赖 better-sidebar 版本

需要 better-sidebar `>= 0.6.0`（已移除内置 office viewer 的版本），否则会与内置 `docx/xlsx/pptx` 注册冲突（`already registered`）。

## 开发

```powershell
pnpm install
pnpm run typecheck   # 类型门禁（需先构建 DSH-better-sidebar 的 lib/types）
pnpm test            # vitest（xlsx→Univer 转换 + 注册描述符）
pnpm run build       # tsdown 双产物（lib/index.js + lib/client.js）
```

> `tsdown.config.ts` 含 `jszip` / `xlsx` 的 browser-entry alias（SheetJS/JSZip 在 CJS 降级后残留 Node builtin 引用）与 `import.meta.resolve` 空定义（pptx-renderer 的 PDF.js 探测）。改动这些库版本时需一并验证 client bundle 不含 Node builtin。

## 安装（profile）

```powershell
# 从 npm 安装（推荐）：
dsh plugin --profile web add @huanlin/dsh-plugin-better-sidebar-plugin-office

# 本地开发（link: 热更新）
dsh plugin --profile web add "link:D:/Projects/deepseek-harness/dsh-better-sidebar-plugin-office"
```

安装后重启 `dsh web` 进程并浏览器硬刷新（`Ctrl+Shift+R`）。

## 运行

- 打开侧边栏 → 在文件资源管理器中打开任意 `.docx` / `.xlsx` / `.pptx`，预览器自动命中。
- Side card 设置页 →「文件预览」清单中可见 `docx / xlsx / pptx` 三张卡片，可单独启用/禁用。

## 检查

| 检查项 | 命令 | 预期 |
|--------|------|------|
| 类型门禁 | `pnpm run typecheck` | 0 错误 |
| 单元测试 | `pnpm test` | 全部通过 |
| 构建 | `pnpm run build` | `lib/client.js` 生成，`window.__ModuleLoader__.load({ id: '@huanlin/dsh-plugin-better-sidebar-plugin-office', ... })` 包裹 |
| bundle 纯度 | 产物中搜索 `node:` 或 `require("fs")` | 不应出现 Node builtin 引用 |
| profile 可见 | `Test-Path ~/.dsh/profiles/web/node_modules/@huanlin/dsh-plugin-better-sidebar-plugin-office/lib/client.js` | True |
| 插件加载 | 重启 `dsh web` + 硬刷新后侧边栏打开 Office 文件 | 预览正常渲染 |
