---
marp: true
theme: academic
paginate: true
---

<!-- _class: card-grid -->

# 実験条件

<div class="cg-container">

<div class="cg-card">
  <span class="cg-title">条件 A: 短系列</span>
  <span class="cg-body">系列長 512。標準的な NLU タスク。BERT-base と同等設定。</span>
</div>

<div class="cg-card">
  <span class="cg-title">条件 B: 中系列</span>
  <span class="cg-body">系列長 2048。文書分類・要約タスク。GPT-2 と同等設定。</span>
</div>

<div class="cg-card">
  <span class="cg-title">条件 C: 長系列</span>
  <span class="cg-body">系列長 8192。長距離依存タスク。Longformer と同等設定。</span>
</div>

<div class="cg-card">
  <span class="cg-title">条件 D: 超長系列</span>
  <span class="cg-body">系列長 16384。書籍レベル。本研究の主要ターゲット。</span>
</div>

</div>
