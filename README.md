# Office Diff Lab

在浏览器本地比较 Word、Excel、PowerPoint 和带文字层的 PDF，查看字符、数值、样式和结构变化，并导出带高亮的原格式副本。文件不会上传，无需后端服务。

## 开发与验证

使用 Node.js 24 和 pnpm，在项目根目录执行：

```sh
cd demo
pnpm install --frozen-lockfile
pnpm run build
node ../scripts/package_office_skill.mjs
pnpm test
pnpm run dev
```

打开 `http://127.0.0.1:4173/`。Skill 打包命令生成测试所需的归档；修改引擎源码后，运行测试前应重新打包。

## 静态部署

构建产物为 `demo/out/`，可部署到静态 HTTP/HTTPS 站点。Vercel 使用以下配置：

- Root Directory：`demo`
- Framework Preset：`Other`
- Install Command：`pnpm install --frozen-lockfile`
- Build Command：`pnpm run build`
- Output Directory：`out`
- Node.js：`24.x`

## 目录

- `demo/`：前端、WASM 源码、构建脚本和测试，详见 [应用说明](demo/README.md)。
- `docs/`、`schemas/`、`examples/`：差异协议文档、机器定义和示例。
- `scripts/`：协议验证与 Skill 打包工具。
- `skills/office-diff/`：Office 差异比较 Skill 及其运行资源。

依赖目录、缓存、构建产物、测试输出、本地部署绑定和环境变量文件由 `.gitignore` 排除。
