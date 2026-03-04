# 访谈纪要融合脚本使用说明（Agent-First）

脚本路径：`scripts/interview_fusion.py`

## 1. 环境准备
1. Python 3.9+
2. 已安装 `python-docx`
3. 准备输入文件（默认同目录）：
   - `转写.docx`
   - `访谈纪要.docx`

## 2. 推荐流程（Codex / Claude Code）
### Step A: 生成 Agent 任务包
```bash
python3 scripts/interview_fusion.py prepare-agent \
  --transcript-docx 转写.docx \
  --summary-docx 访谈纪要.docx \
  --output-dir work \
  --batch-size 25
```

如需每次先清理旧中间文件：
```bash
python3 scripts/interview_fusion.py prepare-agent \
  --transcript-docx 转写.docx \
  --summary-docx 访谈纪要.docx \
  --output-dir work \
  --batch-size 25 \
  --clean
```

产物：
- `work/temp_transcript.md`
- `work/temp_summary.md`
- `work/transcript_chunks.json`
- `work/questions.json`
- `work/llm_input.json`
- `work/agent_tasks/manifest.json`
- `work/agent_tasks/system_prompt.txt`
- `work/agent_tasks/worker_prompt_template.md`
- `work/agent_tasks/batches/Bxxx.input.json`（每批输入）
- `work/agent_tasks/AGENT_TASK.md`

### Step B: 用 Codex / Claude Code 处理批次
按 `work/agent_tasks/manifest.json`：
1. 读取每个 `Bxxx.input.json`
2. 按 `system_prompt.txt` + `worker_prompt_template.md` 生成结果
3. 写回到对应 `Bxxx.output.json`
4. 若识别到纪要外临时问题，可写入 `new_items`（需证据支持）

### Step C: 合并结果并回写 DOCX
```bash
python3 scripts/interview_fusion.py finalize-agent \
  --output-dir work \
  --summary-docx 访谈纪要.docx \
  --output-docx 访谈纪要完成版ClaudeQwen.docx
```

`finalize-agent` 会自动执行：
1. `merge-agent-results`（把所有 `Bxxx.output.json` 合成 `work/llm_results.json`）
2. `apply`（回写 `docx`）

## 3. 仅合并（可选）
```bash
python3 scripts/interview_fusion.py merge-agent-results \
  --manifest-json work/agent_tasks/manifest.json \
  --output-json work/llm_results.json
```

## 4. 回写规则（已内置）
1. 新增内容标红。
2. 已有回复内容仅“改动片段”标红。
3. `confidence >= 75`：正常写入。
4. `55 <= confidence < 75`：写入并自动追加 `（注：低置信度，需要人工审核）`。
5. `confidence < 55`：留空。
6. `new_items` 中状态为 `filled` 的临时问题，会插入“资料清单”前并标红回写。

## 5. 常用参数
- `--project-id`：项目标识（可选）
- `--batch-size`：每批问题数量（默认 25）
- `--strict`：合并时若缺少批次输出则报错（默认缺失批次按 blank 处理）

## 6. API 直连流程（保留）
如果你不走 Codex/Claude 的批次处理，也可以直接调用模型 API：
```bash
python3 scripts/interview_fusion.py pipeline \
  --transcript-docx 转写.docx \
  --summary-docx 访谈纪要.docx \
  --output-dir work \
  --output-docx 访谈纪要完成版ClaudeQwen.docx \
  --batch-size 25
```

## 7. 故障排查
1. 某批次未生成 `Bxxx.output.json`：`finalize-agent` 默认按 blank 补齐；如需强校验，加 `--strict`。
2. 批次输出不是合法 JSON：先修复该批次输出格式，再重新执行 `finalize-agent`。
3. 某些问题未被填充：检查 `work/llm_results.json` 中该 `question_id` 的 `status/confidence/evidence`。
