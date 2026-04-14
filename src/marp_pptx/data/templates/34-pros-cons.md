---
marp: true
theme: academic
paginate: true
---

<!-- _class: pros-cons -->

# 提案手法の利点と制約

<div class="pc-pros">
<li>計算量を 60% 削減</li>
<li>既存パイプラインへのドロップイン置換が可能</li>
<li>理論的な収束保証あり</li>
<li>GPU 並列化に対応</li>
</div>

<div class="pc-cons">
<li>短系列タスクではオーバーヘッドが相対的に大きい</li>
<li>マスク生成器の追加学習が必要</li>
<li>動的マスクの解釈性が限定的</li>
</div>
