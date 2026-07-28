# LLM Wiki Schema

本文件是知识库的核心配置。它定义了 Wiki 的结构、约定和工作流程。作为 LLM，你应当遵循本文件来维护知识库。

---

## 知识库结构

```
kb/
├── CLAUDE.md            # 本文件 — Schema 配置
├── index.md             # 内容目录（所有页面的索引）
├── log.md               # 按时间顺序的操作日志
├── templates/           # Obsidian 模板（用于快速创建标准页面）
│   └── README.md        # 模板使用说明
├── raw/                 # 原始资料（只读，永不修改）
│   └── assets/          # 图片/附件
└── wiki/                # LLM 生成的 Wiki 页面
    ├── entities/        # 实体页面（人物、组织、作品、地点等）
    ├── concepts/        # 概念页面（抽象主题、术语、思想）
song    └── summaries/       # 摘要页面（单篇来源的提炼）
```

## 页面约定

### 文件名
- 使用小写英文 + 连字符，如 `machine-learning.md`
- 如果内容涉及中文专名，可用拼音或英文翻译作为文件名
- 避免空格和特殊字符

### YAML Frontmatter
每个 Wiki 页面必须包含以下 frontmatter：

```yaml
---
title: 页面标题
type: entity | concept | summary
created: 2026-06-05
updated: 2026-06-05
tags: [tag1, tag2]
sources: []  # 涉及到的源文件路径
aliases: []  # 别名/同义词
---
```

### 交叉引用
- 使用 `[[页面名]]` 格式引用其他 Wiki 页面（无 `.md` 后缀）
- 在文中首次提到重要概念时，若对应页面已存在或应当存在，添加链接

### 写作风格
- 使用中文撰写内容，术语首次出现时标注英文原文
- 保持客观、结构化，多用标题分层
- 每个页面开头用 1-2 句话概括核心内容

---

## 工作流程

### 1. Ingest（摄入）

当用户将新资料放入 `raw/` 目录或提供 URL/内容时：

1. **读取原始资料** — 完整阅读源文件
2. **讨论与提炼** — 与用户讨论核心要点，确认哪些信息更重要
3. **创建摘要页面** — 在 `wiki/summaries/` 下创建对应摘要，文件名格式：`YYYY-MM-DD-short-slug.md`
4. **更新实体和概念页面** — 查找并更新 `wiki/entities/` 和 `wiki/concepts/` 中受影响的页面；如果是新知识，创建新页面
5. **更新 index.md** — 添加新页面的条目
6. **更新 log.md** — 追加一条摄入记录
7. **检查矛盾** — 如果新资料与已有内容矛盾，在相关页面中标注并通知用户

### 2. Query（查询）

当用户提问时：

1. **查看 index.md** — 了解所有可用页面
2. **定位相关页面** — 根据问题搜索 `wiki/` 下可能相关的页面
3. **阅读并综合** — 读取相关页面内容，综合回答
4. **引用来源** — 回答中标注引用到的 Wiki 页面名
5. **可选：归档回答** — 如果回答有价值，创建新的 Wiki 页面归档（如 `wiki/concepts/answer-topic.md`）

### 3. Lint（健康检查）

定期（或应要求）执行：

1. **检查矛盾** — 跨页面对比，查找陈述不一致
2. **检查过时内容** — 旧页面是否有已被新来源推翻的论断
3. **查找孤儿页面** — 没有入站链接的页面（`index.md` 中的条目不算）
4. **发现遗漏** — 页面中提到的概念/实体是否应当有自己的页面
5. **查找缺失引用** — 明显应该交叉引用的地方
6. **建议新方向** — 基于现有知识空缺，建议用户寻找哪些新来源

---

## index.md 格式

按类别组织，每页一行：

```markdown
## Entities
- [页面名](wiki/entities/page.md) — 一句话描述

## Concepts
- [页面名](wiki/concepts/page.md) — 一句话描述

## Summaries
- [页面名](wiki/summaries/page.md) — 一句话描述
```

每次 Ingest 后必须更新。

## log.md 格式

每条记录使用统一前缀，方便用 grep 检索：

```markdown
## [2026-06-05] ingest | 资料标题
- 创建了: summary-page, entity-page
- 更新了: concept-page
```

```markdown
## [2026-06-05] query | 用户问题
- 查阅了: page1, page2
```

```markdown
## [2026-06-05] lint | 健康检查
- 发现 2 个孤儿页面
- 建议创建: missing-concept
```

---

## 工具

- **Obsidian** — 用于浏览和编辑 Wiki（注意：路径使用 Windows 风格）
- **Marp** — 用于从 Wiki 内容生成幻灯片
- **git** — 整个 `kb/` 是一个 git 仓库，LLM 可通过 `git diff` 查看变更

---

*本 Schema 是动态的，随着使用不断进化。当发现更好的约定或工作流程时，更新本文件并记录到 log.md 中。*
