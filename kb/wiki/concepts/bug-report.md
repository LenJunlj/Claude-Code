---
title: 缺陷报告
type: concept
created: 2026-06-05
updated: 2026-06-05
tags: [testing, bug, defect, quality-assurance]
aliases: [bug report, 缺陷报告, Bug单, 问题单]
---

# 缺陷报告

**缺陷报告**（Bug Report / Defect Report）是测试执行阶段的主要输出物之一，用于记录和跟踪软件中发现的缺陷。

## 典型组成

- **缺陷ID**：唯一编号
- **标题**：简明扼要描述缺陷
- **严重程度**：Critical / Major / Minor / Trivial
- **优先级**：P0（立即修复）~ P3（可延后）
- **复现步骤**：详细的操作路径
- **实际结果**：当前表现
- **预期结果**：正确的表现
- **环境信息**：操作系统、浏览器、版本号等
- **附件**：截图、日志、视频等
- **状态**：New → Open → Fixed → Verified → Closed

## 相关概念

- [[software-testing-process|软件测试流程]] — 缺陷报告在测试执行阶段产出
- [[test-case|测试用例]] — 执行测试用例时发现缺陷
