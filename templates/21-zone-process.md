---
marp: true
theme: academic
paginate: true
---

<!-- _class: zone-process -->

# 実験手順

<div class="zp-container">

<div class="zp-step">
  <span class="zp-num">1</span>
  <span class="zp-title">データ収集</span>
  <span class="zp-body">公開データセットから10万件のサンプルを取得し、ノイズ除去を実施。</span>
</div>

<div class="zp-step">
  <span class="zp-num">2</span>
  <span class="zp-title">前処理</span>
  <span class="zp-body">正規化、欠損値補完、特徴量エンジニアリングを適用。</span>
</div>

<div class="zp-step">
  <span class="zp-num">3</span>
  <span class="zp-title">モデル学習</span>
  <span class="zp-body">5-fold交差検証でハイパーパラメータを最適化。学習率スケジューリング導入。</span>
</div>

<div class="zp-step">
  <span class="zp-num">4</span>
  <span class="zp-title">評価・比較</span>
  <span class="zp-body">テストセットでAUC、F1スコアを算出し、3つのベースラインと比較。</span>
</div>

</div>
