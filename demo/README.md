# Office Diff Lab — Office / PDF 纯前端 + WebAssembly

Word、Excel、PowerPoint 和带文字层的 PDF 在浏览器本地完成提取、对象匹配、字符差异、预览和原格式高亮导出。交付目录 `out/` 包含 HTML、CSS、JavaScript、WASM、PDF 字体与 CMap 资源、协议与示例文件，可放在任意静态 HTTP/HTTPS 站点。没有后端 API、.NET、Blazor、托管程序集或 C/C# 源码依赖。文件不会上传。

## 本地运行

安装依赖后：

```sh
node tools/build.mjs
node --test tools/*.test.mjs
node serve.mjs
```

打开 `http://127.0.0.1:4173/`。也可以使用 `./build.ps1` 完成构建和检查。依赖由 `pnpm-lock.yaml` 锁定，用 `pnpm install --frozen-lockfile` 安装。Node 和包管理器仅用于开发、构建、测试与本地静态服务，部署后的页面不依赖 Node。

本项目使用原生浏览器 ES modules、DOMParser、XMLSerializer、Web Crypto、Intl.Segmenter、WebAssembly 和本地打包的 fflate。无需 CDN；生产站点只请求同站静态资源。通过 HTTP/HTTPS 使用，直接双击 `file://` 不支持模块加载和 Web Crypto。

## 运行链路

```text
Office ZIP → JavaScript OOXML 包与关系解析
           → Word / Excel / PowerPoint 三个独立提取器
           → 对象匹配 + 独立 WASM 字符对齐
           → ChangeSet + 原文预览
           → JavaScript 修改原包中的指定 XML part → 高亮 Office 副本
```

| 文件 | 职责 |
| --- | --- |
| `browser/wwwroot/engine/ooxml.js` | ZIP 限额、OPC 关系、命名空间、XML 路径、文本投影 |
| `browser/wwwroot/engine/adapters.js` | DOCX 段落、XLSX 单元格、PPTX 空间画布提取 |
| `browser/wwwroot/engine/office-engine.js` | 语义匹配、统一 Change 协议 |
| `browser/wwwroot/engine/character-diff.js` | Unicode grapheme 分词、UTF16 偏移、WASM 调用 |
| `wasm/character-diff.wat` | 手写 WAT 双向 Myers 线性内存搜索 |
| `browser/wwwroot/character-diff.wasm` | 可直接部署的独立算法二进制，零运行时导入 |
| `browser/wwwroot/engine/annotation.js` | 原格式字符高亮与元素标记 |
| `browser/wwwroot/text-diff-view.js` | 字符范围的 HTML 对照呈现 |
| `browser/wwwroot/engine/pdf-engine.js` | PDF.js 提取、页面渲染、PDF 高亮副本导出 |
| `browser/wwwroot/engine/pdf-model.js` | 页面与文字匹配、样式范围比较、页序分析 |
| `browser/wwwroot/pdf-review.js` | PDF 原页双侧预览与区域定位 |

WASM 真实执行 Myers 中点搜索，不是占位模块。JavaScript 负责 grapheme 编码及对齐任务调度。构建时使用 wabt 将 WAT 转成 WASM，不需要 C/C# 编译器、.NET SDK、WASI 或其他语言项目。wabt 与测试用 xmldom 只在开发环境使用；浏览器不加载它们。

## 原文对照与 Excel 大表浏览

原文对照的新旧版本各自支持横向、纵向滚动。点击右侧变更明细会切换到对应页面或工作表，分别定位并选中新旧版本中的原始锚点；插入行导致地址变化时，使用各自的单元格地址。上方支持上一处／下一处变更导航。

Excel 默认使用「变更聚焦」：按工作表展示变更数量，保留变化单元格的相邻行列与首行、首列，折叠其余未变化的数据，可点击分段展开。「完整数据」可分页浏览整个已提取的数据范围。两种模式每侧每页最多渲染 80 行、24 列，点击明细自动打开目标所在页，不再截断在前 60 行、20 列。表头与行号固定，公式变化同时显示公式。新增或删除没有对应锚点的一侧会明确提示。

## 字符差异与导出

字符级差异仅用于两侧均为文本的变化和公式文本。例如文本 `120 → 150` 只高亮 `2 → 5`，同段多处修改分别标记；编号等数字字符串保留前导零并按文本比较。插入、删除、替换保留新旧范围；emoji 和组合字符按 grapheme 比较，锚点单位为 UTF16。

真正的数值（包括公式保存的数值结果）按完整值比较与高亮，不生成字符差异扩展，也不显示字符修改处数。`120 → 150` 展示完整旧值、新值；`1` 与 `1.00` 等数值相等的写法不产生数值事件。比较和展示保留精确十进制字符串，千分位仅通过字符串格式化添加，避免 JavaScript Number 精度或舍入隐藏变化。

Change v1 的 `extensions["ppt2wiki.textDiff"]` 保存字符子差异，三套专属定位器继续绑定各自文件快照。导出前校验 SHA-256 与字符范围，防止旧锚点误写新文件。

协议另有正式的 ChangeSet 级 `extensions["ppt2wiki.impact"]`，定义依赖、范围、影响原因、计算证据和单独的分析覆盖状态，详见 [依赖与影响协议](../docs/office-impact-protocol.md)。构建分发其 Schema。Excel 比较显式启用公式缓存值比较；公式文本不变、已保存结果变化时，产生独立 value 事件并高亮结果。没有缓存时不将其当作空值进行比较。

### 计算影响

页面在实际结果变化的单元格上显示 `ƒx`，右侧可筛选“计算结果”。“计算影响”入口单独展示已变化、未变或待确认的分析记录；相同或缺失缓存的结果显示问号，不能被解读为结果未变。额外重算分析不增加原文件变化数，也不覆盖表格的文件原值。

自动生成的计算影响只纳入已保存结果变化、公式修改，或有静态依赖路径连接到实际变化输入的单元格。单纯因引用未完整解析，不会将无已知变化关联的公式计入影响列表；这些缺口在“分析限制”中按公式单元格去重计数，并保留在导出协议的依赖覆盖信息中。有变化关联但缓存相同或缺失的公式仍标为“结果待确认”。

选择计算结果后，明细提供前后保存值、公式是否变化、候选来源和后续影响。可从结果定位输入，也可从输入定位下游，跨工作表时自动切换。变更聚焦会包含关联输入和必要的中间对象；“查看引用范围”只标记当前公式直接引用的区域，按需显示，保持大范围分页与折叠。

`engine/formula-impact.js` 解析普通 A1、绝对引用、矩形、整行整列和基本跨表引用；字符串中的地址、科学计数常量和外部工作簿目标不会被当作本地单元格。命名、结构化、动态、共享公式等未完整解析的情况会报告缺口。当前没有计算引擎或控制回放，所有自动生成的来源判断为 supported/potential，页面明确说明“原因待确认”。

新增的“公式与连锁变化”示例由 `tools/formula-fixture.mjs` 在构建时生成，覆盖范围合计、连锁计算、跨表引用、公式同时修改及缓存相同/缺失。导出的高亮文件继续只标记真实核心差异，保留原公式及缓存；额外影响提示在页面和 ChangeSet 中查看。

Word 和 PPT 支持普通文本字符高亮，跨 run 拆分时保留原有文字属性；Excel 字符串转为当前单元格的富文本，避免修改共享字符串影响其他单元格。Excel 数值与公式保留原生类型，原文件导出采用单元格高亮；页面明细展示完整数值，公式文本继续提供字符对比。文本以外变化采用元素或单元格标记。

原始 ZIP 中未修改的 part 保留原有载荷字节；重新打包不保证 ZIP 容器逐字节一致。高亮是增加标记的副本，不是接受／拒绝修订机制。

## 示例与边界

### PDF 首版

打开 `/?sample=pdf-budget`，或上传同格式的旧版和新版 PDF。示例包含数字文字修改、字号、颜色、混合片段粗体、文字位置、文字新增、整页增删及调序。左右展示 PDF 原始页面，页面列表使用“旧 1 ↔ 新 3”表明匹配关系；点击明细分别定位各自快照，缺失页面的一侧明确提示。

PDF.js 6.4.299 使用兼容构建，字体、CMaps、解码 WASM 与 worker 全部由本地构建复制，生产无需 CDN。PDF 资源按需加载。PDF-LIB 1.17.1 用于导出；开发测试另用原生 canvas，部署产物不包含原生 Node 模块。静态服务需将 `.mjs` 作为 JavaScript MIME、`.wasm` 作为 WebAssembly MIME 提供，`serve.mjs` 已配置。

PDF 通过独立的 `../schemas/pdf-change.v1.schema.json` 定义 `format: pdf` 和 `profile: pdf-text-v1`，保留 Office v1 的原约束。事件继续使用 before/after、operation/aspect、match、coverage 与 `ppt2wiki.textDiff`。PDF source 不含 OOXML part；locator 使用零基 `pageIndex`、`coordinateSpace: pageViewport` 和 `rect: [x,y,width,height]`，坐标为应用 CropBox、页面旋转与 UserUnit 后的 scale=1 viewport（左上角原点，单位为物理 pt）。文字锚点另含 PDF.js 的 `itemIndices`、提取文字和从 0 开始的 UTF16 range。定位仅对文件摘要和当前提取器有效，文本值不等同于原始 PDF 字节。

匹配先采用作用域内唯一文字锚点，再采用无近似竞争候选的双向文字相似度。重复文字按提取顺序匹配并报告不确定性；唯一不变文字可跨已匹配页面对应。以已匹配页面的相对序列判断调序，插页引起的被动页码偏移不计为移动。整页增删不重复报告所有子文字。文字位置以 0.75 pt 为移动阈值；字体大小和内容自然引起的包围框变化不额外计为移动。文字和样式的不同属性分别记录。

样式按 UTF16 文字范围保存，比较匹配文字中的字体、字号、粗体、斜体和可靠映射的填充颜色，不取首个片段代表整行。无法对应到绘制字符的颜色保持未知并记录 coverage。PDF 数字属于提取文字，按字符比较。明细展示字符变化，原页及导出采用整条文字区域高亮，不声称逐字形精确标记。

导出先检查原文件 SHA-256，再重新提取、核对页面边界、文字、item 索引及区域坐标。输出叠加半透明区域标记，保留文字、页面、裁剪和旋转；整页变化使用边框和浅色标记。这是带标记的副本，不是 PDF 原生修订或可接受/拒绝的编辑记录。

当前未比较图片、图表、线条、背景、批注、透明度等全部外观，也未接入视觉 diff 或 OCR。没有文字层的文件明确拒绝；混合文件中无文字层的页面记录覆盖缺口。多栏、竖排、复杂变换、跨行重排和重复内容的身份匹配仍有局限。每个 PDF 上限 15 MB、150 页、50000 个文字 item、500000 个 UTF16 单元，单页最多 3000 个 item；待匹配候选组合超过 100000 时明确报错。比较结果始终提供覆盖说明，未检测到变化不代表未比较的外观一致。

`tools/pdf-engine.test.mjs` 覆盖真实 PDF 的文字、混合样式、页序、整页增删、重复行、计算限额、真实中文 CID 字体/CMap、90/180/270 度旋转与裁剪、快照和锚点篡改拒绝、原文保留及实际页面渲染。示例和高亮输出也经过 Poppler 渲染检查。

协议校验：`python ../scripts/validate_change_protocol.py artifacts/pdf-budget.changes.json`，与 Office 使用同一入口并按 format 选择 Schema。先按 `../scripts/requirements-change-protocol.txt` 安装验证依赖。

- Word：季度与金额两处字符修改、加粗、段落新增。
- Excel：插入部门行后的语义匹配、金额、公式、负责人单字修改。
- Excel 公式：结果变化标签、来源双向定位、范围提示、跨表传播及待确认结果。
- PPT：文本、表格字符、形状位置、增删、幻灯片调序。

结构预览不是 Office 排版引擎。复杂字段、主题继承、图表、SmartArt、动画等仍属于部分覆盖，页面会说明范围。动态字段内部的字符高亮会明确报错。单文件上限 15 MB，解压上限 100 MB、5000 个条目；不支持 ZIP64、加密包、Strict OOXML 命名空间。计算超出演示上限会报错，不静默降级为粗粒度差异。

构建会清理固定的 `out/` 后重新生成，确保旧 `_framework`、DLL、PDB 不会混入部署包。测试覆盖 WASM 与独立 LCS 的最短编辑一致性、Unicode、三格式示例、快照保护、跨 run 属性和共享字符串隔离。
