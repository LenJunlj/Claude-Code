---
title: LLM Wiki
type: concept
created: 2026-06-05
updated: 2026-06-05
tags: [methodology, knowledge-management, personal-knowledge-base]
sources: [raw/LLM-Wiki-idea.md]
aliases: [LLM 维基, LLM Wiki 模式]
---

# LLM Wiki

**LLM Wiki** 是一种利用 LLM（大型语言模型）增量构建和维护个人知识库的方法论模式。其核心思想是：LLM 不仅从原始资料中检索信息，而是持续构建一个结构化的、相互链接的 Markdown 文件 Wiki。

## 核心理念

与传统的 RAG（检索增强生成）不同——后者每次查询从头检索和综合信息，没有知识积累——LLM Wiki 模式让 LLM 像一个勤勉的图书管理员，在每次摄入新资料时主动更新、整合和优化已有的知识库。

关键区别：
- **RAG**：每次查询重新发现知识
- **LLM Wiki**：知识被编译一次并持续维护

## 三层架构

| 层级 | 内容 | 所有者 |
|------|------|--------|
| **Raw Sources** | 原始资料（文章、论文、数据文件）——只读不可变 | 用户 |
| **The Wiki** | LLM 生成的 Markdown 文件——摘要、实体页、概念页 | LLM |
| **The Schema** | 配置文件（如 [[../../../CLAUDE.md|CLAUDE.md]]）——定义结构和工作流程 | 用户 + LLM 共同进化 |

## 三大操作

- **Ingest（摄入）**：阅读新资料 → 讨论提炼 → 创建/更新 Wiki 页面 → 维护 index 和 log
- **Query（查询）**：搜索 Wiki → 综合回答 → 有价值的回答归档回 Wiki
- **Lint（健康检查）**：查找矛盾、孤儿页面、缺失引用、过时内容

## 为什么有效

知识库维护的瓶颈从来不是阅读或思考，而是"记账工作"——更新交叉引用、保持摘要最新、维护跨页面一致性。人类放弃 Wiki 是因为维护负担增长快于价值增长。LLM 不会厌倦、不会忘记更新交叉引用，一次可以处理 15 个文件。维护成本趋近于零，Wiki 因此得以持续。

## 精神传承

这一思想与 Vannevar Bush 在 1945 年提出的 **Memex**（记忆扩展器）理念一脉相承——一个私人的、精心策展的知识存储，文档之间有关联路径。Bush 无法解决的问题是"谁来维护"，而 LLM 恰好补上了这一环。

## 相关页面

- [[../../../CLAUDE.md|Schema 配置]] — 本 Wiki 的具体实现规范
- [[../../../index.md|知识库索引]] — 所有页面的目录
- [[../../../log.md|操作日志]] — 时间线记录
