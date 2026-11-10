# 计划骨架 · copy.md

- 模板: `templates/【西财蓝1】演示文稿模板.pptx`
- 页数: 10

## 映射结果

```
P 1 · cover          → clone 1  (列 2, 槽 6) ⚠ 标题 25 字 > 样张容量 6 字 ×1.5，建议压缩标题
P 2 · toc            → clone 5  (列 3, 槽 7) ok
P 3 · section        → clone 2  (列 1, 槽 2) ok
P 4 · content        → clone 6  (列 3, 槽 7) ⚠ 4 条内容 > 样张 3 个槽位：建议改表格页（content-table）、拆成两页，或换列数更多的样张；禁止合并
P 5 · content        → clone 7  (列 3, 槽 10) ⚠ 标题 20 字 > 样张容量 12 字 ×1.5，建议压缩标题
P 6 · content-table  → clone 7  (列 3, 槽 10) ⚠ 表格页：确认 removeShapes 已删掉全部装饰，保留标题栏与校徽
P 7 · content        → clone 6  (列 3, 槽 7) ⚠ 4 条内容 > 样张 3 个槽位：建议改表格页（content-table）、拆成两页，或换列数更多的样张；禁止合并; 标题 22 字 > 样张容量 12 字 ×1.5，建议压缩标题
P 8 · content        → clone 7  (列 3, 槽 10) ⚠ 4 条内容 > 样张 3 个槽位：建议改表格页（content-table）、拆成两页，或换列数更多的样张；禁止合并; 标题 22 字 > 样张容量 12 字 ×1.5，建议压缩标题
P 9 · content-table  → clone 7  (列 3, 槽 10) ⚠ 表格页：确认 removeShapes 已删掉全部装饰，保留标题栏与校徽
P10 · end            → clone 13 (列 2, 槽 5) ⚠ 标题 11 字 > 样张容量 3 字 ×1.5，建议压缩标题
```

## 每页槽位

| 页 | clone | 槽位（shapeId:类型:容量字） |
|---|---|---|
| P1 | 1 | 32:title:6, 8:cardTitle:5, 34:cardBody:6, 33:cardBody:12, 2:cardBody:7, 3:cardBody:12 |
| P2 | 5 | 16:title:8, 28:cardTitle:6, 29:cardBody:6, 2:cardTitle:6, 20:cardBody:6, 4:cardTitle:6, 25:cardBody:6 |
| P3 | 2 | 16:title:6, 6:cardBody:2 |
| P4 | 6 | 2:title:12, 9:cardTitle:5, 8:cardBody:44, 11:cardTitle:5, 10:cardBody:44, 13:cardTitle:5, 12:cardBody:45 |
| P5 | 7 | 8:title:12, 5:decor:2, 4:cardTitle:5, 3:cardBody:50, 10:decor:2, 7:cardTitle:5, 6:cardBody:50, 14:decor:2, 12:cardTitle:5, 11:cardBody:37 |
| P6 | 7 | 8:title:12, 5:decor:2, 4:cardTitle:5, 3:cardBody:50, 10:decor:2, 7:cardTitle:5, 6:cardBody:50, 14:decor:2, 12:cardTitle:5, 11:cardBody:37 |
| P7 | 6 | 2:title:12, 9:cardTitle:5, 8:cardBody:44, 11:cardTitle:5, 10:cardBody:44, 13:cardTitle:5, 12:cardBody:45 |
| P8 | 7 | 8:title:12, 5:decor:2, 4:cardTitle:5, 3:cardBody:50, 10:decor:2, 7:cardTitle:5, 6:cardBody:50, 14:decor:2, 12:cardTitle:5, 11:cardBody:37 |
| P9 | 7 | 8:title:12, 5:decor:2, 4:cardTitle:5, 3:cardBody:50, 10:decor:2, 7:cardTitle:5, 6:cardBody:50, 14:decor:2, 12:cardTitle:5, 11:cardBody:37 |
| P10 | 13 | 32:title:3, 34:cardTitle:7, 33:cardBody:12, 2:cardBody:7, 3:cardBody:12 |

## 下一步

1. 把每页 `content` 的文字填进 `texts`（按 slots 的 kind 对应 title/cardTitle/cardBody）；
2. 表格页填 `addShapes[0].data`，确认 `removeShapes`；
3. 先 `build_deck.py --only <页>` 出单页样张（确认门 3），再批量生成。