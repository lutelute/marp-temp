---
marp: true
theme: academic
paginate: true
math: katex
---

<!-- _class: title -->
<!-- _paginate: false -->

# テンプレートギャラリー
## 研究発表スライド用 Marp テンプレート一覧

各スライド種別のサンプルを収録
テンプレート名はタイトルに明記

v1.0 — 2026年 4月

---

<!-- _class: agenda -->

# [Agenda] 目次テンプレート

<div class="agenda-list">

1. 背景・課題の提示
2. 提案手法の説明
3. 実験設定と結果
4. 考察・議論
5. まとめと今後の展望

</div>

---

<!-- _class: divider -->
<!-- _paginate: false -->

# [Divider] セクション区切り

## ここにセクション副題を入れる

---

# [Default] 本文テンプレート

## 問題意識

- 第1の背景情報をここに記述する
- 第2の背景情報をここに記述する
- 既存手法の課題を明確に示す

## 本研究の貢献

<div class="box-accent">

1. 貢献 1: 精度を維持したまま計算量を **60% 削減**
2. 貢献 2: 理論的な保証（収束性、計算量上界）
3. 貢献 3: 3つのベンチマークで SOTA を達成

</div>

---

<!-- _class: rq -->

# [RQ] 研究課題テンプレート

<div class="rq-main">
ここに研究課題（Research Question）を1文で記述する。$O(n^2)$ のような数式も使える。
</div>

<div class="rq-sub">
仮説: ここに仮説を補足的に記述する。
</div>

---

<!-- _class: equation -->

# [Equation] 数式テンプレート

<div class="eq-main">

$$\text{SparseAttn}(Q,K,V) = \text{softmax}\!\left(\frac{QK^\top \odot M}{\sqrt{d_k}}\right)V$$

</div>

<div class="eq-desc">
  <span class="sym">$Q, K, V$</span>
  <span>クエリ・キー・バリュー行列</span>
  <span class="sym">$M$</span>
  <span>学習可能なスパースマスク</span>
  <span class="sym">$\sqrt{d_k}$</span>
  <span>スケーリング係数</span>
  <span class="sym">$\odot$</span>
  <span>要素積（Hadamard product）</span>
</div>

---

<!-- _class: equations -->

# [Equations] 複数数式テンプレート（最適化問題）

<div class="eq-system">
  <div class="row">
    <span class="label">minimize</span>

$$\mathcal{L}(M) = \tfrac{1}{N}\sum_{i=1}^{N} \ell\bigl(f_M(x_i), y_i\bigr) + \lambda \lVert M \rVert_{1}$$

  </div>
  <div class="row">
    <span class="label">subject to</span>

$$M_{ij} \in \{0, 1\}, \quad \forall (i, j)$$

  </div>
  <div class="row">
    <span class="label"></span>

$$\sum_{j} M_{ij} \le k, \quad \forall i$$

  </div>
</div>

<div class="eq-desc">
  <span class="sym">$M$</span>
  <span>スパースマスク行列</span>
  <span class="sym">$\lambda$</span>
  <span>$L_1$ 正則化係数</span>
  <span class="sym">$k$</span>
  <span>各行のアクティブ要素数上限</span>
</div>

---

<!-- _class: figure -->

# [Figure] 図版テンプレート

![w:750](assets/architecture.svg)

<div class="caption"><span class="fig-num">Fig. 1.</span> ここに図のキャプションを記述する。簡潔に、図の読み取り方を示す。</div>

---

<!-- _class: sandwich -->

# [Sandwich] 上下挟み + 3カラムテンプレート

<div class="top">

<p class="lead">上部にリード文を配置。全体の文脈を1〜2文で説明し、下のカラムの読み方を示す。</p>

</div>

<div class="columns c3">
<div>

### カラム A

- 項目 1-1
- 項目 1-2
- 項目 1-3

</div>
<div>

### カラム B

- 項目 2-1
- 項目 2-2
- 項目 2-3

</div>
<div>

### カラム C

- 項目 3-1
- 項目 3-2
- 項目 3-3

</div>
</div>

<div class="bottom">

<div class="conclusion">

**まとめ**: 下部に結論ボックスを配置。3カラムの要点を1文でまとめる。

</div>

</div>

---

<!-- _class: sandwich -->

# [Sandwich+Box] 中央背景ボックス付きテンプレート

<div class="top">

<p class="lead">中央領域に背景付きボックスを配置する例。概念図や図解の代わりに、構造化されたテキストブロックを使う。</p>

</div>

<div class="columns c2">
<div>

<div class="box">

**入力**: 生テキスト系列 $x_1, x_2, \ldots, x_n$

**処理 1**: Embedding → ベクトル列に変換

**処理 2**: Sparse Mask を動的生成

**処理 3**: マスク付き Attention 計算

**出力**: コンテキスト付きベクトル列

</div>

</div>
<div>

### ポイント

- 各処理は独立に差し替え可能
- Flash Attention と併用可能
- マスク生成のオーバーヘッドは 5% 以下
- 入出力の次元は不変 → ドロップイン置換

</div>
</div>

<div class="bottom">

<div class="box-accent">

**設計原則**: モジュラー構成により、既存パイプラインへの導入コストを最小化する。

</div>

</div>

---

<!-- _class: cols-2 -->

# [Cols-2] 2カラムテンプレート

<div class="columns">
<div>

### 左カラム

ここに図、グラフ、またはテキストを配置する。

<div class="box">

数値結果やハイライトをボックスで囲む例:

- 精度: **89.4%**
- 速度: **2.4x**
- メモリ: **-40%**

</div>

</div>
<div>

### 右カラム

図の解説や追加の分析結果をここに記述する。

- 左の結果は3回の試行の平均値
- 標準偏差は ±0.3% 以内
- ベースラインとの差は統計的に有意 ($p < 0.01$)
- 長系列タスクほど改善幅が大きい

</div>
</div>

---

<!-- _class: cols-3 -->

# [Cols-3] 3カラムテンプレート

<div class="columns">
<div>

### ベンチマーク A

<div class="box-accent">

- スコア: **17.9**
- 計算: 0.40x
- 改善: -2.2%

</div>

長系列言語モデリングでの性能。

</div>
<div>

### ベンチマーク B

<div class="box-accent">

- スコア: **89.4**
- 計算: 0.38x
- 改善: +1.1%

</div>

自然言語理解タスクでの性能。

</div>
<div>

### ベンチマーク C

<div class="box-accent">

- スコア: **82.1**
- 計算: 0.35x
- 改善: +3.6%

</div>

長距離依存タスクでの性能。

</div>
</div>

---

<!-- _class: table-slide -->

# [Table] 表テンプレート

## Table 1. 手法間の定量比較

| 手法 | Perplexity ↓ | 計算時間 | メモリ | パラメータ数 |
|------|:-----------:|:------:|:----:|:---------:|
| Transformer | 18.3 | 1.00x | 1.00x | 125M |
| Linformer | 19.1 | 0.52x | 0.48x | 125M |
| Flash Attention | 18.3 | 0.65x | 0.55x | 125M |
| **Ours** | **17.9** | **0.40x** | **0.42x** | **127M** |

<div class="box-accent">

**Ours** が全指標で最良。パラメータ増加はマスク生成器分のわずか 2M。

</div>

<div class="footnote">同一ハードウェア (NVIDIA A100 80GB) で測定</div>

---

<!-- _class: timeline-h -->

# [Timeline-H] 横タイムラインテンプレート

<div class="tl-h-container">

<div class="tl-h-item">
  <div class="tl-h-block">
    <span class="tl-h-year">2017</span>
    <span class="tl-h-text">Transformer</span>
    <div class="tl-h-detail">Self-attention<br>$O(n^2)$</div>
  </div>
</div>

<div class="tl-h-item">
  <div class="tl-h-block">
    <span class="tl-h-year">2020</span>
    <span class="tl-h-text">Linformer</span>
    <div class="tl-h-detail">線形近似<br>$O(n)$</div>
  </div>
</div>

<div class="tl-h-item">
  <div class="tl-h-block">
    <span class="tl-h-year">2022</span>
    <span class="tl-h-text">Flash Attention</span>
    <div class="tl-h-detail">IO最適化</div>
  </div>
</div>

<div class="tl-h-item highlight">
  <div class="tl-h-block">
    <span class="tl-h-year">2026</span>
    <span class="tl-h-text bold">本研究</span>
    <div class="tl-h-detail">Sparse + 収束保証<br>$O(n^{3/2})$</div>
  </div>
</div>

</div>

---

<!-- _class: zone-flow -->

# [Zone-Flow] フローテンプレート

<div class="zf-container">

<div class="zf-box">
  <span class="zf-label">Step 1</span>
  <span class="zf-body">データ収集と前処理。ノイズ除去、正規化、トークナイズ。</span>
</div>

<div class="zf-box">
  <span class="zf-label">Step 2</span>
  <span class="zf-body">特徴量設計とモデル構築。ハイパーパラメータ探索。</span>
</div>

<div class="zf-box">
  <span class="zf-label">Step 3</span>
  <span class="zf-body">評価とベースライン比較。統計検定で有意差確認。</span>
</div>

</div>

---

<!-- _class: zone-compare -->

# [Zone-Compare] 比較テンプレート

<div class="zc-container">

<div class="zc-left">
  <span class="zc-label">従来手法</span>
  <span class="zc-body">全注意計算。計算量 $O(n^2)$。精度は高いがスケーラビリティに課題。大規模データでは実用困難。</span>
</div>

<div class="zc-vs">VS</div>

<div class="zc-right">
  <span class="zc-label">提案手法</span>
  <span class="zc-body">動的スパース注意。$O(n\sqrt{n})$ で同等精度。GPU並列化対応。系列長 16K でも実用速度。</span>
</div>

</div>

---

<!-- _class: zone-matrix -->

# [Zone-Matrix] 2x2 マトリクステンプレート

<div class="zm-container">

<div class="zm-ylabel">精度</div>
<div class="zm-xlabel">計算コスト</div>

<div class="zm-cell zm-tl">
  <span class="zm-label">理想（提案手法）</span>
  <span class="zm-body">高精度 + 低コスト。動的スパースで両立。</span>
</div>

<div class="zm-cell zm-tr">
  <span class="zm-label">力技</span>
  <span class="zm-body">高精度だが高コスト。Full Attention 等。</span>
</div>

<div class="zm-cell zm-bl">
  <span class="zm-label">ベースライン</span>
  <span class="zm-body">低コスト・低精度。ルールベース手法。</span>
</div>

<div class="zm-cell zm-br">
  <span class="zm-label">非効率</span>
  <span class="zm-body">高コスト + 低精度。設計上の問題。</span>
</div>

</div>

---

<!-- _class: zone-process -->

# [Zone-Process] 手順テンプレート

<div class="zp-container">

<div class="zp-step">
  <span class="zp-num">1</span>
  <span class="zp-title">データ準備</span>
  <span class="zp-body">公開データセットから 10 万件を取得。ノイズ除去と BPE トークナイズ。</span>
</div>

<div class="zp-step">
  <span class="zp-num">2</span>
  <span class="zp-title">モデル構築</span>
  <span class="zp-body">12 層 Sparse Attention Block。隠れ次元 768、ヘッド数 12。</span>
</div>

<div class="zp-step">
  <span class="zp-num">3</span>
  <span class="zp-title">学習</span>
  <span class="zp-body">AdamW, lr=3e-4, cosine schedule。100 epoch, batch 64。</span>
</div>

<div class="zp-step">
  <span class="zp-num">4</span>
  <span class="zp-title">評価</span>
  <span class="zp-body">Perplexity, FLOPs, wall-clock time で 3 ベースラインと比較。</span>
</div>

</div>

---

<!-- _class: summary -->

# [Summary] まとめテンプレート（Q&A 表示用）

<ol class="summary-points">
<li>貢献 1: ここに第一の成果を記述する</li>
<li>貢献 2: ここに第二の成果を記述する</li>
<li>貢献 3: ここに第三の成果を記述する</li>
<li>展望: ここに今後の方向性を記述する</li>
</ol>

---

<!-- _class: references -->
<!-- _paginate: false -->

# [References] 参考文献テンプレート

<ol>
<li>
  <span class="author">著者 A et al.</span>
  <span class="title">"論文タイトル 1."</span>
  <span class="venue">会議名, 年.</span>
</li>
<li>
  <span class="author">著者 B et al.</span>
  <span class="title">"論文タイトル 2."</span>
  <span class="venue">ジャーナル名, 年.</span>
</li>
<li>
  <span class="author">著者 C et al.</span>
  <span class="title">"論文タイトル 3."</span>
  <span class="venue">arXiv:XXXX.XXXXX, 年.</span>
</li>
</ol>

---

<!-- _class: appendix -->

# [Appendix] 付録テンプレート

<span class="appendix-label">Appendix A</span>

| パラメータ | 値 | 備考 |
|---|---|---|
| 隠れ次元 | 768 | BERT-base と同一 |
| ヘッド数 | 12 | Multi-head Attention |
| 層数 | 12 | |
| 学習率 | 3e-4 | cosine schedule |
| バッチサイズ | 64 | V100 x 4 |

---

<!-- _class: end -->
<!-- _paginate: false -->

# [End] 終了テンプレート

Questions?

name@university.ac.jp
