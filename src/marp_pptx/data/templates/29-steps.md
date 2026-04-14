---
marp: true
theme: academic
paginate: true
---

<!-- _class: steps -->

# 感情ベクトルの抽出手順

<div class="st-container">

<div class="st-step">
  <span class="st-num">1</span>
  <span class="st-title">感情語の選定</span>
  <span class="st-body">171種の感情語を選定し、各感情に対応するプロンプトを設計</span>
</div>

<div class="st-step">
  <span class="st-num">2</span>
  <span class="st-title">ストーリー生成</span>
  <span class="st-body">各感情について1200本のストーリーを LLM で生成</span>
</div>

<div class="st-step">
  <span class="st-num">3</span>
  <span class="st-title">残差ストリーム取得</span>
  <span class="st-body">生成過程の中間層から残差ストリームベクトルを抽出</span>
</div>

<div class="st-step">
  <span class="st-num">4</span>
  <span class="st-title">ノイズ除去</span>
  <span class="st-body">平均ベクトルと交絡因子の影響を project out して精製</span>
</div>

</div>
