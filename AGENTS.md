# Project Agent Guide

本文件与另一份代理说明文件（`AGENTS.md` / `CLAUDE.md`）保持同步更新。

## 项目目标
- 将 `转写.docx` 的访谈内容融合到 `访谈纪要.docx` 的“回复：”位置。
- 输出 `访谈纪要完成版ClaudeQwen.docx`。

## 核心规则
- 语义匹配优先，关键词仅辅助。
- 同一问题多段回答可合并为 2-4 句。
- 无明确回答时留空。
- 新增内容标红；旧内容仅改动片段标红。
- 低置信度阈值：
  - `confidence >= 75`：正常写入
  - `55 <= confidence < 75`：写入并追加 `（注：低置信度，需要人工审核）`
  - `confidence < 55`：留空
- 支持临时新增问题：可在 `new_items` 输出，最终回写到纪要补充区。

## 推荐流程（Agent-First）
1. 生成任务包
```bash
python3 scripts/interview_fusion.py prepare-agent \
  --transcript-docx 转写.docx \
  --summary-docx 访谈纪要.docx \
  --output-dir work \
  --batch-size 25 \
  --clean
```

2. 按批次处理
- 读取 `work/agent_tasks/manifest.json`
- 对每个 `batches/Bxxx.input.json` 生成 `batches/Bxxx.output.json`
- 遵循：
  - `work/agent_tasks/system_prompt.txt`
  - `work/agent_tasks/worker_prompt_template.md`

3. 合并并回写
```bash
python3 scripts/interview_fusion.py finalize-agent \
  --output-dir work \
  --summary-docx 访谈纪要.docx \
  --output-docx 访谈纪要完成版ClaudeQwen.docx
```

## 批次输出格式
每个 `Bxxx.output.json` 必须包含：
- `project_id`
- `results`（每个 `question_id` 一条）
- `new_items`（可选）

`results[]` 字段：
- `question_id`
- `status`
- `confidence`
- `evidence`
- `answer_draft`
- `note_required`
- `note_text`

`new_items[]` 字段：
- `question_text`
- `status`
- `confidence`
- `evidence`
- `answer_draft`
- `note_required`
- `note_text`

## 关键文件
- `scripts/interview_fusion.py`
- `readMe.md`
- `history.md`
- `prdV2.md`
- `Plan.md`

## 完成定义
- 所有批次输出文件存在且 JSON 合法。
- `finalize-agent` 成功执行。
- 最终文档生成且回写条目数量与问题数量一致。
