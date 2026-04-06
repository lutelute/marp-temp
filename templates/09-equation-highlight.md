---
marp: true
theme: academic
paginate: true
math: katex
---

<!-- _class: equation -->

# 数式スライド — 領域ハイライト

<div class="eq-main">

$$\text{Attention}(Q, K, V) = \text{softmax}\!\left(\frac{\colorbox{#fff3cd}{$QK^\top$}}{\colorbox{#cce5ff}{$\sqrt{d_k}$}}\right)\colorbox{#fff3cd}{$V$}$$

</div>

<div class="eq-desc">
  <span class="sym"><span class="eq-highlight">$QK^\top$</span></span>
  <span>クエリとキーの内積 → 類似度スコア</span>
  <span class="sym"><span class="eq-highlight-b">$\sqrt{d_k}$</span></span>
  <span>スケーリング係数（勾配消失を防止）</span>
  <span class="sym"><span class="eq-highlight">$V$</span></span>
  <span>バリュー行列 → 最終出力の重み付け対象</span>
  <span class="sym">$\text{softmax}$</span>
  <span>正規化 → 注意重みの確率分布化</span>
</div>

<div class="footnote">Vaswani et al., "Attention Is All You Need," NeurIPS 2017</div>
