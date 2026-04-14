---
marp: true
theme: academic
paginate: true
---

<!-- _class: before-after -->

# 改善効果

<div class="ba-before">
  <span class="ba-label">Before</span>
  <span class="ba-body">Full Attention で全ペアを計算。系列長 4096 で OOM が頻発。学習に 72 時間。</span>
</div>

<div class="ba-after">
  <span class="ba-label">After</span>
  <span class="ba-body">Sparse Attention で計算量 60% 削減。系列長 16384 でも安定動作。学習 30 時間。</span>
</div>
