---
marp: true
theme: academic
paginate: true
math: katex
---

<!-- _class: equation -->

# 数式スライド — 基本

<div class="eq-main">

$$\mathcal{L}(\theta) = -\frac{1}{N}\sum_{i=1}^{N} \left[ y_i \log \hat{y}_i + (1-y_i)\log(1-\hat{y}_i) \right]$$

</div>

<div class="eq-desc">
  <span class="sym">$\mathcal{L}(\theta)$</span>
  <span>損失関数（交差エントロピー）</span>
  <span class="sym">$\theta$</span>
  <span>モデルパラメータ</span>
  <span class="sym">$N$</span>
  <span>サンプル数</span>
  <span class="sym">$y_i$</span>
  <span>正解ラベル（0 or 1）</span>
  <span class="sym">$\hat{y}_i$</span>
  <span>モデルの予測確率</span>
</div>
