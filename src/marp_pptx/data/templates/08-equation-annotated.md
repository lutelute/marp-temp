---
marp: true
theme: academic
paginate: true
math: katex
---

<!-- _class: equation -->

# 数式スライド — アノテーション付き

<div class="eq-main">

$$\hat{y} = \underbrace{\sigma}_{\text{活性化}} \left( \overbrace{W}^{\text{重み}} \cdot \underbrace{x}_{\text{入力}} + \overbrace{b}^{\text{バイアス}} \right)$$

</div>

<div class="eq-desc">
  <span class="sym">$\hat{y}$</span>
  <span>出力（予測値）</span>
  <span class="sym">$\sigma(\cdot)$</span>
  <span>シグモイド活性化関数: $\sigma(z) = \frac{1}{1+e^{-z}}$</span>
  <span class="sym">$W$</span>
  <span>重み行列 $\in \mathbb{R}^{m \times n}$</span>
  <span class="sym">$x$</span>
  <span>入力ベクトル $\in \mathbb{R}^{n}$</span>
  <span class="sym">$b$</span>
  <span>バイアスベクトル $\in \mathbb{R}^{m}$</span>
</div>

<div class="footnote">KaTeX の \underbrace / \overbrace を使って数式内に直接アノテーションを記述</div>
