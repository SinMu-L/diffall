# 协议的关键解释

此技能直接输出项目的 Office Change v1，`schemaVersion: 1.0.0`，`comparison.profile: office-semantic-v1`。没有另一套用模型推测的变化列表。

- `documents.before/after`：文件名和原始字节 SHA-256，名字不是身份。
- `changes`：无执行顺序含义的事实观察。每条有 `id`、`operation`、`aspect`、`target`、`property`、`before`、`after`，更新/移动还有 `match`。
- `before/after` 中的 `value` 保留类型，`source` 指向所属快照。数字使用精确十进制字符串；文本不应被当作数字强制转换。
- `target.entityId` 只在本次比较中稳定；`target.label` 是展示描述，不是定位器。
- Word 定位器包含 OOXML part、元素路径；Excel 包含工作表 ID、地址；PPT 包含幻灯片 ID、形状路径及可选表格单元格。
- `extensions["ppt2wiki.textDiff"]`：grapheme 对齐、UTF16 偏移的字符范围。Python 字符索引不能直接替代 UTF16 偏移。
- `extensions["ppt2wiki.impact"]`：Excel 的静态引用与潜在连锁影响；不代表已完成重算或确认因果。缓存值变化和公式文本变化可以是独立事实。
- `coverage`：本次实际覆盖情况。`partial` 与空变化数组不能证明全部文档等价。公式计算结果相同但缓存未知时不能宣称“不受影响”。

ChangeSet 不是 JSON Patch，不包含重建文件需要的全部数据。它不能直接执行合并、恢复旧版或逐条接受/拒绝修改。

完整机器定义位于 `schemas/office-change.v1.schema.json`、`schemas/office-text-diff.v1.schema.json`、`schemas/office-impact.v1.schema.json`，按需读取即可。
