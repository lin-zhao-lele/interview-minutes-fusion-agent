# 项目历史记录（2026-03-04）

## 1. 项目目标
- 将 `转写.docx` 的访谈对话整合到 `访谈纪要.docx` 的“回复：”位置。
- 产出最终文档：`访谈纪要完成版ClaudeQwen.docx`。
- 保持原文档格式与风格，新增/改动内容按规则标红。

## 2. 已确认业务规则（本次会话达成）
- 新增内容标红。
- 改动过的旧内容也要标红，但粒度为“仅改动片段标红”（不是整句整段）。
- 同一问题在转写中有多段回答时，合并成 2-4 句精炼回复。
- 找不到明确对应回答的问题，留空。
- 语义匹配优先，关键词仅作辅助。
- 低置信度项不中断自动化流程，在回复后追加：`（注：低置信度，需要人工审核）`。
- 低置信度阈值：
  - `confidence >= 75`：写入
  - `55 <= confidence < 75`：写入并加注
  - `confidence < 55`：留空
- 新增特征规则（后续补充）：
  - 纪要多个问题可能对应转写中的一次提问，需“拆要点后分别匹配”。
  - 转写中可能有纪要外临时加问，需连同回复一起写回纪要。

## 3. 文档与方案沉淀
- 原始需求：`prd.txt`
- 升级需求文档：`prdV2.md`
- LLM 批处理提示词模板：`llm_prompt_template.md`
- 流程说明（自然语言）：`Plan.md`
- 使用说明（Agent-First）：`interview_fusion_usage.md`
- 总体 README：`readMe.md`

## 4. 工具化实现（已落地）
- 主脚本：`scripts/interview_fusion.py`
- 支持命令：
  - `prepare`
  - `run-llm`
  - `apply`
  - `pipeline`
  - `agent-pack`
  - `merge-agent-results`
  - `prepare-agent`
  - `finalize-agent`
- 新增参数：
  - `prepare-agent --clean`：执行前自动清理 `output-dir`（带安全保护，只允许工作区内路径）
- 新增能力（针对匹配质量优化）：
  - `questions.json / llm_input.json` 追加问题上下文字段：
    - `section_title`
    - `subsection_title`
    - `slot_index_for_question`
    - `slot_total_for_question`
  - `run-llm` / `merge-agent-results` 支持并汇总 `new_items`（临时新增问题）。
  - `apply` 支持回写 `new_items`：在“资料清单”前插入“临时问题 + 回复”并标红。

## 5. Agent-First 流程（Codex / Claude Code）
1. `prepare-agent` 生成任务包（`work/agent_tasks`）
2. Agent 按 `manifest.json` 逐批处理 `Bxxx.input.json`，输出 `Bxxx.output.json`
3. `finalize-agent` 自动合并并回写 docx

## 6. 本次实际执行结果
- 已按 `work/agent_tasks/manifest.json` 处理全部批次（B001~B003）。
- 输出文件：
  - `work/agent_tasks/batches/B001.output.json`
  - `work/agent_tasks/batches/B002.output.json`
  - `work/agent_tasks/batches/B003.output.json`
- 合并与回写命令已执行：
  - `python3 scripts/interview_fusion.py finalize-agent --output-dir work --summary-docx 访谈纪要.docx --output-docx 访谈纪要完成版ClaudeQwen.docx`
- 统计结果：
  - 总问题数：58
  - 已填充：50
  - 留空：8
  - 低置信度加注：21
- 最终文档：
  - `访谈纪要完成版ClaudeQwen.docx`

## 7. 匹配机制摘要（当前实现）
- 纪要侧：以“回复：”为锚点回溯问题，形成 `question_id -> reply_paragraph_index`。
- 纪要侧增强：同题多槽位会记录槽位序号与章节上下文，供模型做差异化匹配。
- 转写侧：按段切块（`chunk_id`）提供证据池。
- 语义匹配：由 LLM 按系统规则完成“问题-证据-答案”对齐，并支持“一次提问拆分匹配多个纪要问题”。
- 临时问题：模型可输出 `new_items`，由脚本合并去重并回写到纪要。
- 后处理：脚本统一做阈值归一、低置信度加注、回写标红。

## 8. 下次讨论可直接延续的点
- 是否对低置信度条目做二次自动重写（仅针对 flagged 项）。
- 是否输出质量报告（每题证据摘要、留空原因、冲突说明）。
- 是否支持多项目批量调度（统一命令封装）。
- 是否把“临时新增问题”默认插入位置改为“九、其他问题”内部编号续写（当前是插入资料清单前的补充块）。
