---
marp: true
theme: academic
paginate: true
math: katex
---

<!-- _class: equations -->

# 最適化問題の定式化

<div class="eq-system">
  <div class="row">
    <span class="label">minimize</span>

$$f(x) = \tfrac{1}{2} x^\top Q x + c^\top x$$

  </div>
  <div class="row">
    <span class="label">subject to</span>

$$A x \le b$$

  </div>
  <div class="row">
    <span class="label"></span>

$$C x = d$$

  </div>
  <div class="row">
    <span class="label"></span>

$$x \ge 0$$

  </div>
</div>

<div class="eq-desc">
  <span class="sym">$x \in \mathbb{R}^{n}$</span>
  <span>決定変数</span>
  <span class="sym">$Q \in \mathbb{S}^{n}_{+}$</span>
  <span>半正定値行列（二次コスト）</span>
  <span class="sym">$c \in \mathbb{R}^{n}$</span>
  <span>線形コスト係数</span>
  <span class="sym">$A \in \mathbb{R}^{m \times n},\ b \in \mathbb{R}^{m}$</span>
  <span>不等式制約</span>
  <span class="sym">$C \in \mathbb{R}^{p \times n},\ d \in \mathbb{R}^{p}$</span>
  <span>等式制約</span>
</div>

<div class="footnote">$Q \succeq 0$ ならば凸二次計画問題（QP）として扱える。</div>
