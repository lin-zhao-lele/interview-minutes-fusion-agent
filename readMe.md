# 访谈纪要融合自动化流程（Agent-First）

本文档说明如何使用本项目，把 `转写.docx` 与 `访谈纪要.docx` 自动融合为最终交付文档 `访谈纪要完成版ClaudeQwen.docx`。

核心脚本：
- `scripts/interview_fusion.py`

推荐模式：
- 使用 `Claude Code` / `Codex` 批处理每个问题批次（不强依赖脚本内置 API 调用）。

---

## 1. 目标与规则

流程目标：
- 从转写中提取回答，填入纪要中每个“回复：”位置。
- 支持“转写一次提问对应纪要多个问题”的拆分匹配。
- 支持识别“纪要外临时新增问题”，并将问题+回复自动写回纪要。
- 新增内容标红。
- 已有回复仅“改动片段”标红。
- 无明确对应答案则留空。
- 低置信度条目在回复后追加：`（注：低置信度，需要人工审核）`。

判定阈值（内置于脚本）：
- `confidence >= 75`：正常写入。
- `55 <= confidence < 75`：写入并加低置信度注释。
- `confidence < 55`：留空。

---

## 2. 环境准备

要求：
- Python 3.9+
- `python-docx`

安装依赖（如未安装）：
```bash
python3 -m pip install python-docx
```

默认输入文件（放在项目根目录）：
- `转写.docx`
- `访谈纪要.docx`

---

## 3. 一图看流程

1. `prepare-agent`：提取问题、转写切块、生成 Agent 任务包。  
2. `Claude Code`：按批次输入 `Bxxx.input.json` 生成 `Bxxx.output.json`。  
3. `finalize-agent`：合并所有批次输出并回写 DOCX。  

---

## 4. 详细使用（以 Claude Code 为例）

### Step 1：生成任务包

在项目根目录运行：
```bash
python3 scripts/interview_fusion.py prepare-agent \
  --transcript-docx 转写.docx \
  --summary-docx 访谈纪要.docx \
  --output-dir work \
  --batch-size 25
```

执行后会生成 `work/agent_tasks/`，里面有清单、提示词和批次输入文件。

如果你希望每次执行前自动清理上次产物，增加 `--clean`：
```bash
python3 scripts/interview_fusion.py prepare-agent \
  --transcript-docx 转写.docx \
  --summary-docx 访谈纪要.docx \
  --output-dir work \
  --batch-size 25 \
  --clean
```

### Step 2：让 Claude Code 逐批处理

你可以在 Claude Code 中用这类指令：
```text
请读取 work/agent_tasks/manifest.json，按 batches 顺序处理。
对每个 batch：
1) 读取对应 input_file（例如 work/agent_tasks/batches/B001.input.json）
2) 严格遵循 work/agent_tasks/system_prompt.txt 与 work/agent_tasks/worker_prompt_template.md
3) 仅输出合法 JSON
4) 写入对应 output_file（例如 work/agent_tasks/batches/B001.output.json）
请处理完所有批次后告诉我完成情况。
```

如果你希望分批人工确认，可以先让 Claude Code 只处理一个批次：
```text
请只处理 batch B001：
读取 work/agent_tasks/batches/B001.input.json，
按 work/agent_tasks/system_prompt.txt 和 worker_prompt_template.md 输出结果，
写入 work/agent_tasks/batches/B001.output.json。
```

### Step 3：合并并回写最终文档

所有批次完成后执行：
```bash
python3 scripts/interview_fusion.py finalize-agent \
  --output-dir work \
  --summary-docx 访谈纪要.docx \
  --output-docx 访谈纪要完成版ClaudeQwen.docx
```

如果你要求“缺任何批次都报错”：
```bash
python3 scripts/interview_fusion.py finalize-agent \
  --output-dir work \
  --summary-docx 访谈纪要.docx \
  --output-docx 访谈纪要完成版ClaudeQwen.docx \
  --strict
```

---

## 5. 这套流程包含的资源（详细）

### 5.1 输入资源
- `转写.docx`：访谈原始转写。
- `访谈纪要.docx`：待填充纪要模板。

### 5.2 核心程序资源
- `scripts/interview_fusion.py`：主流程脚本。

### 5.3 预处理产物（`work/`）
- `temp_transcript.md`：转写可读版。
- `temp_summary.md`：纪要可读版（含段落与表格文本）。
- `transcript_chunks.json`：转写切块数据（含 `chunk_id`）。
- `questions.json`：纪要问题锚点（含段落索引、章节上下文、同题多槽位信息）。
- `llm_input.json`：完整模型输入（问题 + 转写切块）。

### 5.4 Agent 任务包资源（`work/agent_tasks/`）
- `manifest.json`：批次清单与输入/输出映射关系。
- `system_prompt.txt`：系统规则文本。
- `worker_prompt_template.md`：执行模板，约束输出 JSON 格式。
- `AGENT_TASK.md`：给 Codex/Claude 的任务说明。
- `batches/Bxxx.input.json`：每批输入。
- `batches/Bxxx.output.json`：每批输出（由 Claude Code 生成）。

### 5.5 合并与回写产物
- `llm_results.json`：所有批次合并后的统一结果。
- `访谈纪要完成版ClaudeQwen.docx`：最终交付文档。

---

## 6. 批次输出 JSON 要求

每个 `Bxxx.output.json` 需要是对象，结构如下：
```json
{
  "project_id": "访谈纪要",
  "results": [
    {
      "question_id": "Q0001",
      "status": "filled",
      "confidence": 82,
      "evidence": ["T0003", "T0010"],
      "answer_draft": "......",
      "note_required": false,
      "note_text": ""
    }
  ],
  "new_items": [
    {
      "question_text": "临时新增问题文本",
      "status": "filled",
      "confidence": 73,
      "evidence": ["T0045", "T0063"],
      "answer_draft": "......",
      "note_required": true,
      "note_text": "（注：低置信度，需要人工审核）"
    }
  ]
}
```

约束：
- 每个 `question_id` 必须有一条结果。
- `status=blank` 时 `answer_draft` 必须为空字符串。
- `note_required=true` 时 `note_text` 固定为 `（注：低置信度，需要人工审核）`。
- `new_items` 仅用于“纪要中不存在，但转写中有明确证据”的临时问题。

---

## 7. 可用命令总览

```bash
# 基础流程
python3 scripts/interview_fusion.py prepare
python3 scripts/interview_fusion.py run-llm
python3 scripts/interview_fusion.py apply
python3 scripts/interview_fusion.py pipeline

# Agent-First 流程（推荐）
python3 scripts/interview_fusion.py agent-pack
python3 scripts/interview_fusion.py merge-agent-results
python3 scripts/interview_fusion.py prepare-agent
python3 scripts/interview_fusion.py finalize-agent
```

查看参数帮助：
```bash
python3 scripts/interview_fusion.py --help
python3 scripts/interview_fusion.py prepare-agent --help
python3 scripts/interview_fusion.py finalize-agent --help
```

---

## 8. 常见问题

1. 某些 `Bxxx.output.json` 没生成怎么办？  
默认 `finalize-agent` 会把缺失批次按 `blank` 处理。若要严格控制，使用 `--strict`。

2. 批次输出不是合法 JSON 怎么办？  
修复该批次输出文件后重新执行 `finalize-agent`。

3. 某问题没有被填入回复怎么办？  
检查 `work/llm_results.json` 对应项是否 `status=blank`、`confidence<55` 或缺少证据。

4. 如何批量处理多个项目？  
每个项目使用独立 `--output-dir`（例如 `work_projA`、`work_projB`），避免文件互相覆盖。

5. 转写中有临时加问，怎么进入最终纪要？  
在批次输出中写入 `new_items`，`finalize-agent` 会自动将其插入“资料清单”前并标红。

---

## 9. 建议的 Claude Code 执行顺序

1. 你先在终端运行 `prepare-agent`。  
2. 把 `work/agent_tasks/AGENT_TASK.md` 交给 Claude Code。  
3. Claude Code 处理完所有 `Bxxx.input.json` 并写回 `Bxxx.output.json`。  
4. 你在终端运行 `finalize-agent` 生成最终 DOCX。  

这是一套可复用的标准生产流程，后续同类“转写 + 纪要”项目只需替换输入文档并重复以上步骤。
