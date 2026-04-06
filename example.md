---
marp: true
theme: academic
paginate: true
math: katex
---

<!-- _class: title -->
<!-- _paginate: false -->

# 効率的な Transformer 拡張手法の提案
## — Sparse Attention による計算量削減 —

山田 太郎 $^{1}$, 佐藤 花子 $^{2}$

$^{1}$ XX大学 工学研究科 &emsp; $^{2}$ YY研究所

Conference on Machine Learning 2026 &ensp;|&ensp; 2026年 9月 15日

---

<!-- _class: divider -->
<!-- _paginate: false -->

# 1. 背景と動機

## 大規模モデルの計算コスト問題

---

# 研究背景

## 問題意識

- Transformer モデルの大規模化が進行
- 計算コストが二次的に増大 → 実用上の障壁
- 既存の効率化手法は精度低下を伴うことが多い

## 本研究の貢献

<div class="box-accent">

1. 精度を維持したまま計算量を **60% 削減** する Sparse Attention 機構
2. 理論的な計算量解析と収束保証
3. 3つのベンチマークで SOTA を達成

</div>

---

<!-- _class: timeline-h -->

# 関連研究の流れ

<div class="tl-h-container">

<div class="tl-h-item">
  <span class="tl-h-year">2017</span>
  <span class="tl-h-text">Transformer<br>$O(n^2)$ attention</span>
</div>

<div class="tl-h-item">
  <span class="tl-h-year">2020</span>
  <span class="tl-h-text">Linformer<br>線形近似</span>
</div>

<div class="tl-h-item">
  <span class="tl-h-year">2021</span>
  <span class="tl-h-text">Flash Attention<br>IO最適化</span>
</div>

<div class="tl-h-item">
  <span class="tl-h-year">2023</span>
  <span class="tl-h-text">Ring Attention<br>分散処理</span>
</div>

<div class="tl-h-item highlight">
  <span class="tl-h-year">2026</span>
  <span class="tl-h-text bold">本研究<br>Sparse + 収束保証</span>
</div>

</div>

---

<!-- _class: divider -->
<!-- _paginate: false -->

# 2. 提案手法

## Sparse Attention with Convergence Guarantee

---

<!-- _class: equation -->

# コア数式

<div class="eq-main">

$$\text{SparseAttn}(Q,K,V) = \text{softmax}\!\left(\frac{\colorbox{#fff3cd}{$Q K^\top \odot M$}}{\colorbox{#cce5ff}{$\sqrt{d_k}$}}\right)V$$

</div>

<div class="eq-desc">
  <span class="sym">$Q, K, V$</span>
  <span>クエリ・キー・バリュー行列</span>
  <span class="sym"><span class="eq-highlight">$M$</span></span>
  <span>学習可能なスパースマスク（$M_{ij} \in \{0, 1\}$）</span>
  <span class="sym"><span class="eq-highlight-b">$\sqrt{d_k}$</span></span>
  <span>スケーリング係数</span>
  <span class="sym">$\odot$</span>
  <span>要素積（Hadamard product）</span>
</div>

---

<!-- _class: equation -->

# 計算量解析

<div class="eq-main">

$$T(n) = \underbrace{O(n \cdot k)}_{\text{スパース注意}} + \underbrace{O(n \log n)}_{\text{マスク更新}} \ll \underbrace{O(n^2)}_{\text{標準 Attention}}$$

</div>

<div class="eq-desc">
  <span class="sym">$n$</span>
  <span>系列長</span>
  <span class="sym">$k$</span>
  <span>各トークンが参照する近傍数（$k \ll n$）</span>
  <span class="sym">$T(n)$</span>
  <span>1層あたりの計算時間</span>
</div>

<div class="footnote">$k = O(\sqrt{n})$ と設定することで $T(n) = O(n^{3/2})$ を達成</div>

---

<!-- _class: sandwich -->

# 提案手法のアーキテクチャ

<div class="top">

入力系列 → Embedding → N 層の Sparse Attention Block → タスクヘッド

</div>

<div class="columns c3">
<div>

### Sparse Mask 生成

- 入力に基づき動的に生成
- Top-k 選択 + Gumbel Softmax
- 勾配が伝播可能

</div>
<div>

### Attention 計算

- マスク適用後の疎行列演算
- Flash Attention と併用可能
- メモリ使用量 $O(nk)$

</div>
<div>

### Mask 更新

- $L$ ステップごとに再計算
- 収束保証定理に基づく
- オーバーヘッド: 全体の 5%

</div>
</div>

<div class="bottom">

<div class="box">

**設計指針**: マスク生成→注意計算→マスク更新のサイクルを繰り返し、精度と効率のバランスを自動調整

</div>

</div>

---

<!-- _class: divider -->
<!-- _paginate: false -->

# 3. 実験と結果

## 3つのベンチマークでの評価

---

<!-- _class: table-slide -->

# 定量的結果

## 言語モデリング（WikiText-103）

| 手法 | Perplexity ↓ | 計算時間 | メモリ | パラメータ数 |
|------|:-----------:|:------:|:----:|:---------:|
| Transformer | 18.3 | 1.00x | 1.00x | 125M |
| Linformer | 19.1 | 0.52x | 0.48x | 125M |
| Flash Attention | 18.3 | 0.65x | 0.55x | 125M |
| **Ours** | **17.9** | **0.40x** | **0.42x** | **127M** |

<div class="box-accent">

**最良の Perplexity** かつ **最少の計算時間** を同時に達成。パラメータ増加はわずか 2M（マスク生成器分）。

</div>

---

<!-- _class: cols-2 -->

# 結果の可視化

<div class="columns">
<div>

![w:440](https://via.placeholder.com/440x300/f0f2f5/1a1a2e?text=Learning+Curve)

<div class="small muted center">Fig. 1: 学習曲線の比較</div>

- 提案手法は **50 epoch** で収束
- 標準 Transformer は 120 epoch 必要
- 収束後の安定性も優れる

</div>
<div>

![w:440](https://via.placeholder.com/440x300/f0f2f5/1a1a2e?text=Sparsity+Pattern)

<div class="small muted center">Fig. 2: 学習されたスパースパターン</div>

- 局所的な注意が支配的
- 長距離依存は少数のハブトークンが担当
- タスクに応じてパターンが変化

</div>
</div>

---

<!-- _class: cols-3 -->

# ベンチマーク別結果

<div class="columns">
<div>

### WikiText-103

<div class="box-accent">

- PPL: **17.9**
- 計算: 0.40x
- 改善: -2.2%

</div>

</div>
<div>

### GLUE (avg)

<div class="box-accent">

- Score: **89.4**
- 計算: 0.38x
- 改善: +1.1%

</div>

</div>
<div>

### Long Range Arena

<div class="box-accent">

- Acc: **82.1**
- 計算: 0.35x
- 改善: +3.6%

</div>

</div>
</div>

<div class="footnote">Long Range Arena では長系列タスクが多いため、スパース化の恩恵が最大</div>

---

<!-- _class: divider -->
<!-- _paginate: false -->

# 4. まとめ

---

<!-- _class: sandwich -->

# まとめと今後の展望

<div class="top">

<div class="box-primary">

**本研究の貢献**: 収束保証付き Sparse Attention により、精度を維持しつつ計算量 60% 削減を達成

</div>

</div>

<div class="columns c2">
<div>

## 主要な成果

1. 3ベンチマーク全てで SOTA
2. 計算量 $O(n^{3/2})$ の理論保証
3. 既存手法との併用が容易

</div>
<div>

## 今後の課題

1. マルチモーダルへの拡張
2. さらなるスパース化の限界解析
3. ハードウェア最適化の検討

</div>
</div>

<div class="bottom">

本研究は JST CREST (JPMJCR20D3) の支援を受けた。

</div>

---

<!-- _class: references -->
<!-- _paginate: false -->

# References

<ol>
<li>
  <span class="author">Vaswani, A. et al.</span>
  <span class="title">"Attention Is All You Need."</span>
  <span class="venue">NeurIPS, 2017.</span>
</li>
<li>
  <span class="author">Wang, S. et al.</span>
  <span class="title">"Linformer: Self-Attention with Linear Complexity."</span>
  <span class="venue">arXiv:2006.04768, 2020.</span>
</li>
<li>
  <span class="author">Dao, T. et al.</span>
  <span class="title">"FlashAttention: Fast and Memory-Efficient Exact Attention with IO-Awareness."</span>
  <span class="venue">NeurIPS, 2022.</span>
</li>
<li>
  <span class="author">Liu, H. et al.</span>
  <span class="title">"Ring Attention with Blockwise Transformers for Near-Infinite Context."</span>
  <span class="venue">ICLR, 2024.</span>
</li>
<li>
  <span class="author">Devlin, J. et al.</span>
  <span class="title">"BERT: Pre-training of Deep Bidirectional Transformers."</span>
  <span class="venue">NAACL-HLT, 2019.</span>
</li>
</ol>

---

<!-- _class: end -->
<!-- _paginate: false -->

# Thank you

Questions?

taro.yamada@xx-university.ac.jp
