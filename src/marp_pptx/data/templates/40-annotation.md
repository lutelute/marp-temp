---
marp: true
theme: academic
paginate: true
---

<!-- _class: annotation -->

# 図の注釈

<div class="an-figure">

![w:500](assets/architecture.svg)

</div>

<div class="an-notes">
<li>入力層で BPE トークナイズを適用</li>
<li>中間層でスパースマスクを動的生成</li>
<li>出力層で次トークン予測を実行</li>
<li>マスク生成器は 2M パラメータ</li>
</div>
