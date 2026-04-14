---
marp: true
theme: academic
paginate: true
---

<!-- _class: zone-compare -->

# 手法の比較

<div class="zc-container">

<div class="zc-left">
  <span class="zc-label">従来手法</span>
  <span class="zc-body">バッチ処理ベース。計算量 $O(n^2)$。精度は中程度だが安定性が高い。大規模データではスケーラビリティに課題。</span>
</div>

<div class="zc-vs">VS</div>

<div class="zc-right">
  <span class="zc-label">提案手法</span>
  <span class="zc-body">ストリーム処理対応。計算量 $O(n \log n)$。精度向上しつつリアルタイム処理を実現。GPU 並列化に対応。</span>
</div>

</div>
