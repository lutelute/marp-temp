---
marp: true
theme: academic
paginate: true
---

<!-- _class: figure -->

# 図解説スライド

![w:700](../assets/architecture.svg)

<div class="caption"><span class="fig-num">Fig. 1.</span> 提案手法のアーキテクチャ概要。入力系列を Embedding 後、N 層の Sparse Attention Block で処理し、タスクヘッドから最終出力を生成する。</div>

<div class="description">

- **入力層**: 生データを前処理し、特徴ベクトルに変換
- **中間層**: Sparse Attention Block を N 段積層
- **出力層**: タスク固有のヘッドで最終予測を生成

</div>
