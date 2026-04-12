---
marp: true
theme: academic
paginate: true
---

<!-- _class: zone-matrix -->

# 研究アプローチの分類

<div class="zm-container">

<div class="zm-ylabel">精度</div>
<div class="zm-xlabel">計算コスト</div>

<div class="zm-cell zm-tl">
  <span class="zm-label">理想領域</span>
  <span class="zm-body">高精度 + 低コスト。提案手法が目指すターゲット。</span>
</div>

<div class="zm-cell zm-tr">
  <span class="zm-label">力技</span>
  <span class="zm-body">高精度だが高コスト。大規模モデルや総当たり探索が該当。</span>
</div>

<div class="zm-cell zm-bl">
  <span class="zm-label">ベースライン</span>
  <span class="zm-body">低コスト・低精度。ルールベースや単純なヒューリスティクス。</span>
</div>

<div class="zm-cell zm-br">
  <span class="zm-label">非効率</span>
  <span class="zm-body">高コストなのに低精度。設計上の問題があるアプローチ。</span>
</div>

</div>
