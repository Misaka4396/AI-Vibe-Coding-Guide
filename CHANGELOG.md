# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [v1.0.1] - 2026-09-11

### Fixed

- README 预览图改用 jsDelivr 绝对地址（`https://cdn.jsdelivr.net/gh/...`）。
  原因：仓库相对路径的图片由 `raw.githubusercontent.com` 提供，该域名在大陆网络下被 DNS 污染
  （解析为 0.0.0.0），导致图片在仓库主页不显示；改用绝对地址后 GitHub 会经
  `camo.githubusercontent.com` 代理抓取，该域名在大陆可达。
  已实测：三张预览图均以 869×1228 原始尺寸正常加载。

## [v1.0.0] - 2026-09-11

### Added

- 《AI Vibe-Coding 使用指南》Word 文档（`docs/AI-Vibe-Coding-Guide.docx`），32 页、34 张表格，
  含封面、文档信息、修订记录与带页码的目录。
- 同版 PDF（`docs/AI-Vibe-Coding-Guide.pdf`），由 Microsoft Word 渲染导出，目录页码已刷新。
- **第一章　总览**：三阶段流水线全景图、五条核心原则（规划与执行分离 / 每条 Prompt 自包含 /
  可校验优先 / 小步提交可回滚 / 验证看证据）、工具链定位速查表。
- **第二章　阶段一：用 Chatbox 生成任务 Prompt**：
  - 「代码架构师」System Prompt 全文照录，可直接复制使用；
  - 五步工作法（复述 → 追问 → 拆解 → 排序 → 生成）与子任务卡七要素模板；
  - 五维度拆解框架（数据 / 逻辑 / 接口 / 表现 / 基础设施）；
  - 端到端实战示例：「新闻聚合 + 每日中文 PDF 日报」需求 → 5 条集中追问 →
    T-01…T-08 子任务拆解表 → 依赖 DAG → 2 条完整生产级 Prompt（去重引擎、中文 PDF 渲染）
    → 六维度风险提示与测试策略；
  - 6 类常见失败模式与对策。
- **第三章　阶段二：用 Claude Code 与 DeepSeek Harness 实现 Prompt**：
  - DeepSeek Anthropic 兼容端点（`api.deepseek.com/anthropic`）配置与三步冒烟；
  - `--effort`（low / medium / high / xhigh / max）挡位选择指南；
  - 单任务执行五步循环（任务文件 → 非交互执行 → 交互执行 → 自跑验收 → 人工复核 diff）；
  - git worktree 并行编排方案；
  - 实测 CLI 参数速查表（`-p`、`--effort`、`--model`、`--permission-mode`、
    `--max-budget-usd`、`--output-format`、`--append-system-prompt`、`-c`、`--add-dir`）；
  - DeepSeek Harness（`dsh`）定位、启动方式与成熟度说明；
  - 含 `Refs:` / `Prompt-Version:` footer 的可追溯提交信息规范。
- **第四章　阶段三：用 Hermes 跑一遍测试流程（P1–P11）**：11 项工程规范，
  每项含规范定义（Role / Context / Task / Constraints / Input / Output / Acceptance Criteria）
  + Hermes 执行层（怎么跑、怎么验证）+ 可复制的验证命令；另含验证报告应包含的章节清单。
- **第五章　风险提示**：六大结构性风险（规范过重、工具版本时效、测试反模式、多语言不一致、
  开源与内部侧重不同、门禁过严阻塞交付）+ 六类 AI 编码特有风险（臆造 API、静默失败、
  权限过大、密钥泄漏、上下文漂移、无审查合并）+ 10 项上线前检查清单。
- **附录**：Prompt 七要素速查卡、Hermes 验证命令清单、推荐目录结构、术语表。
- 中英双语说明文档：`README.md`（英文主版）与 `README.zh.md`（简体中文），顶部互相切换。
- `tools/build_guide.py`：文档内容源与 python-docx 排版逻辑，可重新生成 DOCX。
- `assets/`：封面、目录、规范章节三张预览图。

### Verified

- Word COM 打开无修复：1714 段落 / 34 表格 / 正文 17316 词。
- 导出 PDF：32 页，1.30 MB；`TablesOfContents(1).Update()` 后目录页码齐全。
- 文本层校验（pymupdf）：关键词命中正常，U+FFFD / □ 缺字方块 0 处。
- Markdown 标记残留扫描：字面 `**` 仅 3 处、反引号 14 处，全部位于代码块内的
  通配符与代码片段（刻意保留）。

### Notes

- 文档中的 CLI 参数与官方指令均取自 2026-09-11 实测环境
  （Claude Code v2.1.126、`deepseek-ai/deepseek-harness`），并刻意不写死未来版本号。
- DeepSeek Harness 处于开发者预览期，官方提示会有兼容性破坏变更。
