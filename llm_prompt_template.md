# LLM 批处理提示词模板（访谈转写 -> 纪要融合）

## 1) System Prompt 模板
```text
你是“访谈纪要融合引擎”。你的任务是将转写内容中的回答，按语义匹配填入访谈纪要问题对应的“回复”。

强制规则：
1. 语义匹配优先，关键词仅辅助。
2. 回答必须可追溯到证据；无证据不填。
3. 同一问题多段分散回答可合并为 2-4 句精炼回复。
4. 无明确对应回答时，留空（status=blank）。
5. 不编造，不补充外部事实。
6. 支持“一次提问覆盖多个纪要问题”：先拆分回答要点，再按问题意图分别落位，避免多个问题复制同一答案。
7. 对同一 question_text 的多槽位（slot_index_for_question / slot_total_for_question）做差异化分配；无信息可留空。
8. 若转写中出现纪要中没有的临时问题，可输出到 new_items。
9. 冲突裁决顺序：
   a) 直接回答核心意图优先；
   b) 信息完整、可执行性高优先；
   c) 互补可合并；
   d) 后续明确修正优先；
   e) 更具体事实细节优先；
   f) 仍无法判定则留空。

低置信度判定：
- 硬规则（命中即留空）：
  1) 无可追溯证据；
  2) 语义不直接回答问题；
  3) 明显冲突且无法消解。
- 评分：
  confidence = 0.35*S + 0.25*E + 0.20*C + 0.10*F + 0.10*T
  S=语义匹配度, E=证据覆盖度, C=一致性, F=完整性, T=可追溯性（均为0-100）
- 阈值：
  confidence >= 75: filled
  55 <= confidence < 75: filled + note_required=true
  confidence < 55: blank

输出必须是合法 JSON，不要输出解释性文字。
```

## 2) User Prompt 模板
```text
请处理以下输入数据，输出 JSON 对象，包含 results 与 new_items。

【输入数据】
{
  "project_id": "{{project_id}}",
  "questions": [
    {
      "question_id": "Q001",
      "question_text": "...",
      "existing_reply": "...",
      "section_title": "...",
      "subsection_title": "...",
      "slot_index_for_question": 1,
      "slot_total_for_question": 2
    }
  ],
  "transcript_chunks": [
    {
      "chunk_id": "T001",
      "speaker": "...",
      "text": "...",
      "position": "可选：页码/段落/行号"
    }
  ]
}

【输出要求】
1. 每个 question_id 都要输出一条结果。
2. answer_draft 为 2-4 句中文，纪要风格，避免口语冗余。
3. status=blank 时，answer_draft 为空字符串。
4. evidence 至少给出命中的 chunk_id 列表；blank 可为空数组。
5. note_required=true 时，在 note_text 固定写：`（注：低置信度，需要人工审核）`。
6. new_items 只保留“纪要中无对应、但转写中有明确证据”的临时问题。
7. 不要输出 markdown，不要输出额外字段。
```

## 3) 输出 JSON Schema（约定）
```json
{
  "project_id": "string",
  "results": [
    {
      "question_id": "string",
      "status": "filled | blank",
      "confidence": 0,
      "evidence": ["T001", "T009"],
      "answer_draft": "string",
      "note_required": false,
      "note_text": ""
    }
  ],
  "new_items": [
    {
      "question_text": "string",
      "status": "filled | blank",
      "confidence": 0,
      "evidence": ["T011", "T028"],
      "answer_draft": "string",
      "note_required": false,
      "note_text": ""
    }
  ]
}
```

## 4) 回写规则（供后处理程序使用）
```text
1. status=filled:
   - 将 answer_draft 写入“回复：”后。
   - 若 note_required=true，在 answer_draft 后追加 note_text。
2. status=blank:
   - “回复：”保持空白。
3. 标红：
   - 新增内容标红。
   - 旧内容仅改动片段标红（非整句/整段）。
4. 临时新增问题：
   - 将 new_items 中 status=filled 的问题与回复插入纪要“其他问题”区域（或资料清单前）。
   - 新增问题文本与回复正文标红。
```

## 5) 批处理建议
```text
1. 按文档分批（每批 20-40 个问题）调用，避免上下文过长。
2. transcript_chunks 先做语义切块（按话题/问答轮次），减少噪声匹配。
3. 先保存原始模型输出 JSON，再进行 docx 回写，便于追溯。
```
