# 执行环境与资源

此包的执行入口是 `scripts/office_diff.py`。它用 Python 标准库读写当前宿主可见的文件，通过 Pyodide 的 `js.eval` 加载 `assets/office-diff.bundle.js`，并等待原项目 JS/WASM 引擎的 Promise。包内包含 XML DOM 适配器，支持没有浏览器 DOMParser 的 Web Worker，不修改宿主 DOMParser/XMLSerializer。

宿主必须允许 Python 访问技能资源及所选工作区文件，并允许 Pyodide 的 JS 桥接、JavaScript 执行、WebAssembly、Web Crypto 和 Intl.Segmenter。仅提供纯 CPython、禁止 JS 桥接或只导入 SKILL.md 文本的环境无法执行此包。自检错误不能替换成模型自行猜测的 diff。

导入 Weave 时选择包含 SKILL.md 的整个 `office-diff` 文件夹。若收到 ZIP，先解压再选择该文件夹。全部运行资源都是文本文件，WASM 已编码在 bundle 内；无需二进制导入、外部 CDN、MCP、浏览器扩展或 Node 运行时。

## Python 接口

- `await self_test()`：检查运行环境，并验证内置生成的 DOCX/XLSX/PPTX 样例的预期变化和同文件零变化。不读取用户文档。
- `await compare_bytes(before_bytes, after_bytes, before_name, after_name)`：返回完整 ChangeSet 字典，不写文件。
- `await compare_files(before, after, output_dir, limit=12, max_chars=14000)`：读取两个文件，向一个尚不存在的新目录写三个产物，返回可直接打印给 AI 的摘要。
- `query_changes(changes_path, offset=0, limit=12, max_chars=14000, aspect=None, operation=None, change_ids=None)`：按完整 JSON 生成分页摘要。offset 相对于筛选后的变化列表。
- `get_change(changes_path, change_id)`：返回一条完整变化，包括字符差异；大文本应继续使用截断分页，避免一次打印超出工具结果限制。

`limit` 范围为 1–100；`max_chars` 范围为 4000–60000，按 `json.dumps(..., ensure_ascii=False, separators=(",", ":"))` 的字符数限制摘要。每个截断的字符串使用 `{preview, originalChars, truncated: true}` 显式标记，完整值仍在 changes.json。coverage、快照和分类计数始终保留；必要元数据本身超过预算时函数报错，不静默丢弃。

HTML 报告转义文档内容，不含脚本、外链或网络请求。完整 JSON 含文件内已提取的变化值，按工作区文档处理。

## 范围

支持 `.docx`、`.xlsx`、`.pptx`，不支持 PDF、旧二进制 Office、加密包、Strict OOXML、ZIP64。沿用原项目限制：单文件 15 MB、最多 5000 个 ZIP 条目、总声明解压大小 100 MB，字符算法有计算量限制。内存中会有文件字节、Base64 和解压后的数据副本，大文档需在宿主实际设备上验证。

初版仅提供差异 JSON 与独立 HTML 报告，不携带原项目完整应用界面或 PDF.js 资源。公式缓存比较和静态影响扩展沿用原引擎；不重算公式，不验证 Office 分页、字体渲染、溢出、母版或未覆盖对象。

bundle 由仓库 `scripts/package_office_skill.mjs` 从原引擎构建；`assets/engine-manifest.json` 记录源文件与 WASM 摘要。更新原引擎后应重新打包、重新自检。第三方许可在 `assets/THIRD_PARTY_LICENSES.txt`。
