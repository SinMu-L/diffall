# Office Change 协议草案

版本为 `1.0.0`。协议记录两个 Office 文件快照之间已经识别出的语义变化，供差异列表、LLM 解读和原文件高亮共同使用。Word、Excel 和 PPT 各自提取和匹配对象，最终输出相同的 `ChangeSet`。

原始文件负责保真。Change 保存比较所需的值与源文件锚点，不保存完整文档，也不承诺仅凭差异重建文件。编辑与合并需要另行定义 Patch 协议、依赖顺序和冲突处理。

## 文件与验证

机器定义见 `../schemas/office-change.v1.schema.json`，采用 [JSON Schema Draft 2020 12](https://json-schema.org/draft/2020-12)。`../examples/` 中有三类格式的完整示例，摘要均为合成占位值，没有对应的 Office 文件。

协议有三层验证。JSON Schema 检查字段和类型；语义验证检查 ID、作用域、格式与操作约束；适配器结合原文件检查摘要、锚点、字符范围和真实内容。前两层通过后，仍需要第三层才能执行准确的高亮。`../scripts/validate_change_protocol.py` 实现前两层及协议回归用例。

正式的依赖与影响分析见 [office-impact-protocol.md](office-impact-protocol.md)，机器定义为 `../schemas/office-impact.v1.schema.json`。ChangeSet 级 `extensions["ppt2wiki.impact"]` 登记相关对象、前后引用、范围、计算依据、因果证据、多来源传播与拆分合并。扩展独立版本为 `1.0.0`，不改变核心 `changes` 的事实差异含义。

## ChangeSet

一次比较对应一个 ChangeSet，且只能比较同一格式的两个文件。

| 字段 | 约束与含义 |
| --- | --- |
| `schemaVersion` | 固定为 `1.0.0` |
| `id` | 比较结果 ID，由调用方分配 |
| `format` | `docx`、`xlsx` 或 `pptx` |
| `documents.before` 和 `documents.after` | 各自的文件名和原始文件字节 SHA256 |
| `comparison` | 比较配置，包含 profile、比较维度和可选容差 |
| `coverage` | 本次配置范围内的覆盖情况及明确的缺口 |
| `changes` | 无执行顺序含义的 Change 数组 |
| `extensions` | 可选的命名空间扩展，如 `ppt2wiki.review` |

`documents` 的摘要绑定输入文件快照；`name` 仅用于展示。ZIP 重打包也可能改变文件摘要，此时必须重新绑定或计算比较结果。所有锚点都只对其所属快照有效。

`comparison.profile` 固定为 `office-semantic-v1`；下文定义其规则。`aspects` 明确本次请求比较的维度，每条 Change 的 aspect 都必须包含在其中。没有列入的维度属于范围外；列入但无法处理的内容必须报告为 coverage 缺口。

## Change 字段

| 字段 | 约束与含义 |
| --- | --- |
| `id` | 在 ChangeSet 内唯一 |
| `operation` | `add`、`remove`、`update` 或 `move` |
| `aspect` | 变化维度，见下一节 |
| `target.entityId` | 在本次比较内标识同一个语义对象；同对象的多条变化共用此 ID |
| `target.kind` | 如 paragraph、textRange、cell、slide、shape、image |
| `target.label` | 面向用户的位置描述，不用于匹配或定位 |
| `target.subtype` | 可选的格式特定子类型 |
| `property` | 目标规范对象内的 JSON Pointer，整体新增或删除时为空字符串 |
| `before` 和 `after` | 对应快照下的状态，或明确的缺失状态 `null` |
| `match` | update 和 move 必填；记录新旧对象的匹配依据和评分 |
| `groupId` | 可选，用于将相关变化显示为一组；不增加一次操作的含义 |
| `extensions` | 可选的非核心扩展 |

`entityId` 由匹配阶段分配，生命周期为一次 ChangeSet。它不等于 Shape ID、单元格地址或元素路径，也不保证跨多次比较不变。所有 placement 中的父对象 ID 也使用这个比较内的身份空间，即使父对象本身没有变化事件。

`property` 遵循 [RFC 6901](https://www.rfc-editor.org/rfc/rfc6901)，指向适配器的规范语义对象，例如 `/text`、`/style/font/bold`。它不指向 Change 的 JSON，也不指向 OOXML；`~` 与 `/` 在路径片段内分别转义成 `~0`、`~1`。协议与 [JSON Patch](https://www.rfc-editor.org/rfc/rfc6902) 分别定义，不能把 Change 数组交给 JSON Patch 执行器。

## 操作与维度

| operation | aspect | property | before | after |
| --- | --- | --- | --- | --- |
| add | existence | 空字符串 | null | 对象状态 |
| remove | existence | 空字符串 | 对象状态 | null |
| update | text、value、formula、style、geometry、media、metadata | 对应属性路径 | 状态 | 状态 |
| move | order 或 parent | `/placement` | 位置状态 | 位置状态 |

`add` 和 `remove` 表示语义对象整体出现或消失，其值是该对象的规范摘要，类型为 object，不要求包含可重建的全部细节。同一个 entityId 不得同时出现整体增删与属性变化。新增整页幻灯片时，默认只产生整页事件；后代内容作为该页的摘要，不再逐个重复新增。

更新以最小有意义属性为单位。同一个文本框同时改文字、移动位置、加粗，产生三条事件，entityId 相同。`geometry` 固定对应完整 `/geometry`，便于准确表达尺寸与变换；其他样式可以细化到 `/style/font/bold`。

| aspect | property 和值 |
| --- | --- |
| text | `/text`；text 值 |
| value | `/value`；text、number、boolean、null 或 error 值 |
| formula | `/formula`；text 为公式，null 表示对象仍存在但没有公式 |
| style | `/style` 或它的子路径；值由规范样式字段决定 |
| geometry | `/geometry`；EMU 整数坐标、尺寸、旋转和翻转信息 |
| order | `/placement`；相同父对象中的顺序变化 |
| parent | `/placement`；父对象变化，且同时包含新的顺序 |
| media | `/asset`；媒体资源的字节摘要、类型及包内 part |
| metadata | `/metadata` 或它的子路径；明确纳入比较范围的元数据 |

拖动 PPT 文本框属于 update geometry。幻灯片调序、同层图形堆叠顺序变化属于 move order。把元素放进另一组合属于 move parent。父对象变化同时引起顺序变化时只发一条 parent 事件。

placement 为 `{ "parentEntityId": "deck", "index": 2 }`，index 是对应快照中语义兄弟对象序列的零基位置，不是 XML 子节点索引。适配器先匹配对象，再计算顺序变化；插入或删除导致的被动序号偏移不逐个产生 move。可用 LCS 等算法得到一组稳定的顺序变化，事件不声称还原用户真实的编辑步骤。

## 状态与值

状态由 `value` 与 `source` 组成。before 中 source 隐式属于 documents.before；after 中 source 隐式属于 documents.after。状态不再复制 snapshot 字段，避免同一条事件内部出现相互矛盾的快照引用。

```json
{
  "value": { "type": "number", "data": "1500000" },
  "source": {
    "part": "/xl/worksheets/sheet1.xml",
    "locator": { "type": "excel", "sheetId": 1, "address": "C4" }
  }
}
```

| 表达 | 含义 |
| --- | --- |
| `before: null` | 旧版中没有这个语义对象 |
| `before.value: { "type": "null", "data": null }` | 对象存在，所比较属性为空 |
| `before.value: { "type": "text", "data": "" }` | 对象存在，文本为空字符串 |

Value 类型为 text、number、boolean、null、error、object、array、asset。number 的 data 使用十进制字符串，支持指数记法，比较按十进制数值进行，避免 JSON 数字传输改变精度；`1.0` 与 `1` 不产生数值事件。日期默认保留 Excel 原始数值，显示格式另作 style 比较，不猜测为时间戳。

geometry 的 unit 固定为 emu。space 为 slide 时表示幻灯片坐标；为 parent 时表示源格式声明的父坐标框，不代表已经求出屏幕绝对位置；为 inline 时表示流式内联对象，x 与 y 必须为 null，只比较其显式尺寸与变换。Word 相对页边距、段落等定位依据在提取模型中保留，其变化作为 metadata 比较；需要比较这些布局变化时应同时请求 metadata，不能把 Word 未经排版计算的位置伪装成画布坐标。

rotation 保留 OOXML 的整数角度单位，每度 60000。unit、space、x、y、width、height、rotation、flipH、flipV 均必填；默认旋转为 0，默认翻转为 false，由提取器显式补齐。group 的几何状态还必须包含 childFrame，用 x、y、width、height 表示其子坐标框，单位同为 emu。childFrame 参与几何比较；未实现组合变换时报告相应 coverage 缺口。仅父组合变换时不为每个子元素重复生成几何事件。

asset 的摘要比较媒体原始字节，不声称比较其视觉内容。图片文件更换但像素相同仍可能产生媒体事件；相同字节被重命名为另一 media part 时不产生媒体事件。

## 原文件锚点

source.part 为从包根开始的规范 part URI，例如 `/ppt/slides/slide3.xml`。不得包含查询串、片段、反斜杠或点路径段。适配器必须通过包关系解析其上下文，不能假设 `slide3.xml` 等于当前第 3 页。

| locator.type | 主要字段 | 使用规则 |
| --- | --- | --- |
| word | elementPath、可选 paragraphId 和 range | elementPath 精确指向快照内的段落或容器；paragraphId 用于交叉核对 |
| excel | sheetId、可选 address、row、column、elementPath | sheetId 通过 workbook 关系解析；地址只在该快照内有效 |
| powerpoint | slideId、可选 shapePath、cell、range、elementPath | slideId 通过 presentation 关系解析；shapePath 是从外层组合到目标的原生 Shape ID 链 |
| xml | elementPath | 三种格式共用的原生 XML 元素回退锚点，只表示定位能力 |

elementPath 是从 part 的 XML 根元素开始，逐级选取元素子节点的零基索引数组；忽略文本节点、注释和 XML 声明。`[]` 表示根元素。路径在同一快照内精确，不能独立用于跨版本匹配。解析时保留原始 Markup Compatibility 结构，不先展开或重写 AlternateContent。

Excel address 为不带 `$` 的大写 A1 单格地址，实际边界不得超过 XFD1048576。地址标识逻辑单元格，允许底层 XML 没有对应的 cell 元素；此时按工作表上下文读取其空值和有效样式。Excel row、column 为一基；PPT cell 的 row、column 为零基。PPT 合并单元格以实际拥有文本的原生单元格为锚点，跨度单独保留在表格模型中。

Word 和 PPT 的 range 使用零基 UTF16 start 与 length，对应 source 容器的规范文本投影。范围不得切开代理对。Word 提取器拼接可见 `w:t`、将 tab 映射为 `\t`、br 与 cr 映射为 `\n`；不把字段指令当可见文字。PPT 投影在段落间加入 `\n`，a:br 对应 `\n`，字段采用显示结果。两者都保留实际空白，不做 trim、NFC 或 NFKC 转换，并保存字符到原始 XML 节点的映射。

profile 使用 Word 现有修订的 accepted 视图，即包含插入、排除删除、包含 moveTo、排除 moveFrom；提取不实际修改原文件。无法处理的复杂字段、修订或兼容性分支应报告缺口。字段显示结果可参与比较，改字段内容另由格式适配器判断。

## 匹配与比较配置

match.method 为 nativeId、structure、content、composite 或 manual；evidence 记录可以审计的匹配理由。confidence 是适配器给出的 0 到 1 评分，不是统计概率，也不构成允许自动改写文件的授权。原生 ID 只能在其作用域和其他证据支持下使用；同 ID 不保证相同对象，不同 ID 也不意味着对象被删除。

匹配存在歧义时必须记录 ambiguousMatch。提取失败或不支持的对象不得直接当作删除或新增；仍可对有确定锚点的其他对象输出变化。

`office-semantic-v1` 规定以下归一化规则：

- 不把 ZIP 条目顺序、XML 前缀或属性顺序、关系 ID 重编号当作语义变化。关系先解析到实际目标，内容引用不按 rId 字符串比较。
- Word 和 PPT 的文本比较基于规范投影，不因同样的文字被拆成不同 runs 而报变化；样式单独比较。超链接等语义变化应作为 metadata 比较，而不是因文本相同而丢弃。
- 样式比较解析后的有效属性。只有样式索引、styleId 或定义存放位置改变，而有效属性相同时，不产生可见样式事件。无法解析继承和主题时记录 unresolvedStyle。语义相同的有效颜色等用一致表示；样式来源信息保留在提取模型，不作为可见样式变化。
- Excel 解析共享字符串，区分文本、数字、布尔值、错误和空值。未物化的空格与不带其他语义的显式空单元格等价，不因 XML 多出一个空 cell 就报新增。
- Excel 公式采用无前导等号的 OOXML 公式文本。共享公式须展开到目标单元格，失败则报告缺口；不承诺证明两条不同公式数学等价。缓存结果默认不比较，只有显式启用 compareCachedFormulaValues 才作为 value 事件，并注明缓存可能未重算。
- 公式文本相同但文件中保存的结果不同，启用缓存值比较后可以独立产生 value 更新。原因、引用路径和结果新鲜度由 `ppt2wiki.impact` 承载；不得仅凭公式包含某单元格就声称已确认因果。额外重算而未保存于原始快照的值只能作为扩展中的计算观察，不能冒充核心 before/after 的原文件值。
- geometryToleranceEmu 默认 0，作用于坐标、尺寸和 childFrame 的逐分量绝对差。旋转、翻转及坐标系变化始终比较；inline 的 null 位置不参与数值容差计算。未超过容差的变化不输出，但原始精度保留在提取模型中。
- 不从文字自行推导业务等价性。Word 中的“120万”是文本；Excel 中的 1200000 是数值。LLM 可生成共同的解释，但不能覆盖原始类型与事实。

## 覆盖情况与输出

coverage.status 为 complete 或 partial。complete 只表示本次 comparison 范围没有已知缺口，不表示实现了所有 Office 功能。partial 必须列出至少一个 issue，说明 unsupportedFeature、opaqueContent、unresolvedStyle、ambiguousMatch、unresolvedAnchor 或 truncatedContent。

issue.source 只能在 side 为 before 或 after 时提供；跨两侧的问题需省略 source 或拆成两个问题。仅比较 SmartArt 或图表的原始 part 摘要无法代替其内容、数值或布局比较；如请求覆盖这些维度，必须明确记录缺口。`changes: []` 且 partial 只能解释为“已处理范围没有识别到变化”。

前端按 target.kind、aspect 和 source 渲染差异；LLM 可以只接收类型、旧值、新值和 label 的紧凑视图，完整协议继续保留。批注由 Word、Excel、PPT 各自的适配器完成：先核对对应文件摘要，再解析锚点和核对原值，最后生成输出副本。任何路径或值不符均报告定位失败，不自动猜测另一个对象。

新增对象在新版标记，删除对象在旧版标记。新版没有被删对象时，默认提供旧版副本或删除清单；不在新版不存在的位置伪造锚点。Word 修订、Excel 样式高亮和 PPT 叠加标记是不同的渲染策略，颜色等展示策略不放入 Change 核心字段。

## 一致性与演进

字符级差异使用可选扩展 `changes[].extensions["ppt2wiki.textDiff"]`，机器约束见 `../schemas/office-text-diff.v1.schema.json`。父 Change 继续保存完整旧值、新值与对象锚点，扩展记录属性值内部的字符变化，不增加整体变化事件的数量，也不改变现有协议版本。

扩展的 unit 固定为 utf16，granularity 为 grapheme。字符比较以 Unicode 字素簇为单位，避免拆开 emoji、组合音标或零宽连接序列；start 与 length 使用对应旧值或新值字符串的 UTF16 偏移。beforeLength 和 afterLength 是这两个字符串的 UTF16 总长度。旧版和新版分别计数，长度改变不会使后续锚点沿用错误的位置。

每个 edit 包含 operation 和 before、after 两段 `{ start, length, text }`。operation 为 insert、delete 或 replace；新增时 before.length 为 0，删除时 after.length 为 0。两侧区间必须递增、互不重叠，text 必须与原字符串切片完全一致；区间之间的未变化文字应相同，按 edit 替换后必须能重建完整新字符串。同一句中的分散变化分别记录，例如 `甲1乙2丙` 到 `甲3乙4丙` 产生两段替换。

例如 `120 万元` 改为 `150 万元`，只记录 `2` 到 `5`：

```json
{
  "unit": "utf16",
  "granularity": "grapheme",
  "beforeLength": 6,
  "afterLength": 6,
  "edits": [
    {
      "operation": "replace",
      "before": { "start": 1, "length": 1, "text": "2" },
      "after": { "start": 1, "length": 1, "text": "5" }
    }
  ]
}
```

默认仅为 text、文本 value 和公式文本生成此扩展，且前后值必须同为 text。编号等数字字符串继续按文本比较并保留前导零；真正的 number 值（包括公式保存的数值结果）按数值比较，不生成字符差异扩展，展示完整旧值与新值，不显示字符修改处数。数值相等的不同写法（如 `1` 与 `1.00`）不产生数值事件。数值展示保留精确十进制字符串，增加千分位不得经过可能舍入的 Number 转换。

既有数值十进制字符串字符扩展的校验规则保持兼容，默认预览不使用这类字符范围。任何格式化后的偏移均不得用于原文件写回。Excel 数值和公式的原文件标记仍以单元格为单位，避免将数值或公式转换为富文本；页面中的公式文本继续支持字符级对比。Word、PPT 和 Excel 的普通文本可根据字符范围拆分 run，保留未变化文字与原有格式。

同一个 entityId 的 kind 必须一致；同一 entityId、aspect、property 不重复输出。允许同一个对象在同一轮比较中同时有内容更新和移动。一个对象同时发生 parent 与 order 变化时合并为单条 parent 事件。提取模型中记录父子关系，以检查整对象增删与后代重复事件。

两侧值在对应维度的比较规则下相等时不输出 update。数组顺序不表示执行顺序；展示可以按文档位置排序。消费者支持的 schemaVersion 必须显式检查；核心字段变更发布新版本，未知核心字段拒绝，额外元数据放入有命名空间的 extensions。扩展不能改变操作与定位语义。

JSON Schema 能检查结构和本地条件，跨字段比较由语义验证实现。源文件摘要、part 解析、ID 实际存在性、范围边界、组内变换、表格合并和字符映射由 Office 适配器验证；未读取原文件的协议验证器不能证明高亮或渲染正确。
