---
name: thesis-defense-ppt-review
description: Review, polish, and repair thesis defense PowerPoint decks, especially Chinese academic defense PPTs. Use when Codex works on a .pptx for thesis defense, proposal defense, group meeting, or academic presentation and must check structure, literature-review logic, academic wording, slide centering, references, logos, page numbers, picture borders, connector layers, and layout defects.
---

# Thesis Defense PPT Review

Use this skill to turn a thesis-defense PPT into a formal, presentation-ready deck. Prioritize the PPT as a live presentation, not a paper draft.

## Core Rules

- Edit the PPT only. Do not modify thesis source files unless explicitly asked.
- Keep only the final deck when requested. Remove intermediate PPT copies only after confirming the final file path.
- Use thesis language conservatively. Prefer "本研究表明", "初步验证", and "实验结果显示" over overclaiming phrases such as "本研究证明" unless the thesis explicitly uses that claim.
- Avoid presenter/editor notes in slide body text: "这页", "本页", "背景呈现逻辑", "回答什么", "说明", "作用", "后续处理", "怎么展示".
- Convert note-like phrases into presentation phrases:
  - "背景呈现逻辑" -> "综述线索"
  - "回答什么" -> "验证框架"
  - "设计逻辑" -> "实验路径"
  - "后续处理" -> "迭代路径"
  - "案例说明" -> "案例价值"
- Prefer formal academic assertion lines over conversational explanations.
- When adding literature review content, include representative sources, what each research stream concludes, and how the thesis inherits or addresses the gap.

## Structure Checklist

Make thesis chapters visible in the PPT structure:

- 绪论: background, policy, problem, research content, research questions.
- 文献综述: core concepts, representative studies, research gap / literature review.
- 研究设计: research route, theoretical framework, platform architecture, evaluation indicators.
- 平台实现: system architecture, modules, workflows, screenshots, AI/OJ integration.
- 实验验证: design, reliability/validity, quality experiment, difficulty experiment, data figures.
- 应用案例: representative task, teacher side, student side, feedback loop.
- 总结展望: conclusions, innovation points, limitations, future work.

## Visual Rules

- Center all main body content horizontally and vertically within the usable slide area.
- Keep title, logo, footer, page number, and source lines fixed when shifting body content.
- Leave visible breathing room below the title rule. Body blocks should generally start at least 0.35-0.45 inches below the top title line.
- Avoid nested cards and heavy decorative images. Use meaningful diagrams, screenshots, or data figures.
- Remove meaningless generated illustrations. Visuals must serve a claim.
- Keep table/body text readable, generally about 14 pt when space allows.
- Ensure all text stays inside its box.
- Ensure image borders match image bounds; do not leave oversized frames.
- Keep one page number only, bottom right.
- Keep the school logo top right on every non-cover slide unless the template intentionally differs.
- Use a consistent academic style: white background, restrained BNU blue, thin rules, stable spacing.

## Common Defects To Check

- Duplicate page numbers, especially old template numbers plus new bottom-right numbers.
- Duplicate references on the same slide. Keep one bottom source line.
- Top-heavy slides whose body starts too close to the title line.
- Left-heavy slides after removing right-side visuals.
- Connector lines rendered above cards or text. Connectors should sit below nodes or be routed around labels.
- Reused figures without need.
- Image frames larger than the pictures.
- Table font too small or text clipped by cells.
- Slide body language that sounds like a note to the presenter rather than a defense slide.
- Conclusion statements that overclaim relative to the thesis evidence.

## Validation Workflow

1. Inspect slide titles and deck outline.
2. Scan text for note-like phrases and overclaims.
3. Check each slide for body horizontal and vertical centering.
4. Check page numbers, logos, source lines, image borders, and connector layers.
5. Check literature-review slides for source-backed logic: source -> finding -> gap -> thesis response.
6. Run a final overlap and duplicate-reference check.
7. Summarize changes concisely with the final PPT path.

See `references/pptx-python-checks.md` for reusable python-pptx inspection snippets.
