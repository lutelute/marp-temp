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

<!-- _class: overview -->

# [Overview] 概要テンプレート

<div class="ov-lead">上部に要約文を配置。中央に概念図や全体像の図を入れ、下部にキーポイントを箇条書きする。研究の全体像を1枚で伝えるためのスライド。</div>

![w:700](assets/architecture.svg)

<div class="caption">Fig. 1. 全体構成図。ここに図の読み方を記載。</div>

<div class="ov-points">
<li>知見 1: 計算量を $O(n\sqrt{n})$ に削減しつつ精度を維持</li>
<li>知見 2: 動的スパースマスクがタスクに応じた注意パターンを学習</li>
<li>知見 3: 既存の効率化手法と直交する設計で併用可能</li>
</div>

---

<!-- _class: result -->

# [Result] 結果テンプレート

<div class="rs-lead">左に図/グラフ、右に考察を配置する。実験結果の提示と解釈を1枚にまとめるレイアウト。</div>

<div class="rs-figure">

![w:500](assets/learning-curve.svg)

<div class="caption">Fig. 2. 学習曲線の比較</div>

</div>

<div class="rs-analysis">
<li>提案手法は 50 epoch で収束（標準の 2.4 倍速い）</li>
<li>収束後の Perplexity は標準 Transformer と同等</li>
<li>学習初期のスパースマスクは局所的 → 後期にグローバルパターンを獲得</li>
<li>Flash Attention との併用でさらに 15% 高速化</li>
</div>

---

<!-- _class: steps -->

# [Steps] 手順テンプレート（横ステップ）

<div class="st-container">

<div class="st-step">
  <span class="st-num">1</span>
  <span class="st-title">データ準備</span>
  <span class="st-body">データセット取得、前処理、トークナイズ</span>
</div>

<div class="st-step">
  <span class="st-num">2</span>
  <span class="st-title">モデル構築</span>
  <span class="st-body">Sparse Attention Block の実装と組み立て</span>
</div>

<div class="st-step">
  <span class="st-num">3</span>
  <span class="st-title">学習と最適化</span>
  <span class="st-body">交差検証、ハイパーパラメータ探索</span>
</div>

<div class="st-step">
  <span class="st-num">4</span>
  <span class="st-title">評価</span>
  <span class="st-body">ベースライン比較、統計検定</span>
</div>

</div>

---

<!-- _class: quote -->

# [Quote] 引用テンプレート

<div class="qt-text">科学とは、知識の体系ではなく、思考の方法である。</div>

<div class="qt-source">Carl Sagan, The Demon-Haunted World (1995)</div>

---

<!-- _class: history -->

# [History] 年表テンプレート

<div class="hs-container">

<div class="hs-item">
  <span class="hs-year">2017</span>
  <span class="hs-event">Transformer アーキテクチャの提案 (Vaswani et al.)</span>
</div>

<div class="hs-item">
  <span class="hs-year">2019</span>
  <span class="hs-event">BERT による事前学習モデルの普及</span>
</div>

<div class="hs-item">
  <span class="hs-year">2022</span>
  <span class="hs-event">Flash Attention による IO 最適化</span>
</div>

<div class="hs-item">
  <span class="hs-year">2025</span>
  <span class="hs-event">スパース注意機構の理論的保証（本研究）</span>
</div>

</div>

---

<!-- _class: panorama -->

# [Panorama] パノラマテンプレート

<div class="pn-text">左側にリードテキストを配置し、全体像を説明する。右側には概念図やアーキテクチャ図を大きく表示する。テキストと図の組み合わせで、直感的な理解を促す。</div>

![w:600](assets/architecture.svg)

---

<!-- _class: kpi -->

# [KPI] 主要指標テンプレート

<div class="kpi-container">

<div class="kpi-item">
  <span class="kpi-value">89.4%</span>
  <span class="kpi-label">分類精度</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">2.4x</span>
  <span class="kpi-label">処理速度向上</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">-40%</span>
  <span class="kpi-label">メモリ削減</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">127M</span>
  <span class="kpi-label">パラメータ数</span>
</div>

</div>

---

<!-- _class: pros-cons -->

# [Pros-Cons] 利点・制約テンプレート

<div class="pc-pros">
<li>計算量を 60% 削減</li>
<li>既存パイプラインへのドロップイン置換が可能</li>
<li>理論的な収束保証あり</li>
<li>GPU 並列化に対応</li>
</div>

<div class="pc-cons">
<li>短系列タスクではオーバーヘッドが大きい</li>
<li>マスク生成器の追加学習が必要</li>
<li>動的マスクの解釈性が限定的</li>
</div>

---

<!-- _class: definition -->

# [Definition] 定義テンプレート

<div class="df-term">スパースアテンション (Sparse Attention)</div>

<div class="df-body">入力系列の全ペア間ではなく、選択的に注意重みを計算する手法。計算量を $O(n^2)$ から $O(n\sqrt{n})$ 以下に削減しつつ、タスク性能を維持する。</div>

<div class="df-note">関連概念: Self-Attention, Multi-Head Attention, Linear Attention</div>

---

<!-- _class: diagram -->

# [Diagram] 大型図版テンプレート

![w:900](assets/architecture.svg)

<div class="caption">Fig. 1. 提案手法の全体構成。入力から出力までのデータフロー。</div>

---

<!-- _class: gallery-img -->

# [Gallery-Img] 画像ギャラリーテンプレート

<div class="gi-container">

<div class="gi-item">

![w:400](assets/architecture.svg)

<div class="gi-caption">条件 A</div>
</div>

<div class="gi-item">

![w:400](assets/learning-curve.svg)

<div class="gi-caption">条件 B</div>
</div>

<div class="gi-item">

![w:400](assets/architecture.svg)

<div class="gi-caption">条件 C</div>
</div>

<div class="gi-item">

![w:400](assets/learning-curve.svg)

<div class="gi-caption">条件 D</div>
</div>

</div>

---

<!-- _class: highlight -->

# [Highlight] ハイライトテンプレート

<div class="hl-text">動的スパースマスクにより、計算量を 60% 削減しながら精度を維持できる</div>

---

<!-- _class: checklist -->

# [Checklist] チェックリストテンプレート

<div class="cl-container">
<li class="done">データセットの前処理と分割</li>
<li class="done">ベースラインモデルの実装と検証</li>
<li class="done">提案手法の実装</li>
<li>ハイパーパラメータ探索</li>
<li>統計的有意性検定</li>
<li>アブレーション実験</li>
</div>

---

<!-- _class: annotation -->

# [Annotation] 注釈付き図テンプレート

<div class="an-figure">

![w:500](assets/architecture.svg)

</div>

<div class="an-notes">
<li>入力層で BPE トークナイズを適用</li>
<li>中間層でスパースマスクを動的生成</li>
<li>出力層で次トークン予測を実行</li>
<li>マスク生成器は 2M パラメータ</li>
</div>

---

<!-- _class: before-after -->

# [Before-After] 改善前後テンプレート

<div class="ba-before">
  <span class="ba-label">Before</span>
  <span class="ba-body">Full Attention で全ペアを計算。系列長 4096 で OOM が頻発。学習に 72 時間。</span>
</div>

<div class="ba-after">
  <span class="ba-label">After</span>
  <span class="ba-body">Sparse Attention で計算量 60% 削減。系列長 16384 でも安定動作。学習 30 時間。</span>
</div>

---

<!-- _class: funnel -->

# [Funnel] ファネルテンプレート

<div class="fn-container">

<div class="fn-stage">
  <span class="fn-label">Raw Data</span>
  <span class="fn-value">1,000,000 件</span>
</div>

<div class="fn-stage">
  <span class="fn-label">前処理済み</span>
  <span class="fn-value">850,000 件</span>
</div>

<div class="fn-stage">
  <span class="fn-label">品質フィルタ後</span>
  <span class="fn-value">500,000 件</span>
</div>

<div class="fn-stage">
  <span class="fn-label">最終データセット</span>
  <span class="fn-value">100,000 件</span>
</div>

</div>

---

<!-- _class: stack -->

# [Stack] スタックテンプレート

<div class="sk-container">

<div class="sk-layer">
  <span class="sk-name">Application</span>
  <span class="sk-desc">推論 API、バッチ処理、リアルタイム予測</span>
</div>

<div class="sk-layer">
  <span class="sk-name">Model</span>
  <span class="sk-desc">Sparse Attention Block、マスク生成器</span>
</div>

<div class="sk-layer">
  <span class="sk-name">Framework</span>
  <span class="sk-desc">PyTorch 2.0、CUDA カーネル、Mixed Precision</span>
</div>

<div class="sk-layer">
  <span class="sk-name">Infrastructure</span>
  <span class="sk-desc">NVIDIA A100 x 4、NVLink、高速ストレージ</span>
</div>

</div>

---

<!-- _class: card-grid -->

# [Card-Grid] カードグリッドテンプレート

<div class="cg-container">

<div class="cg-card">
  <span class="cg-title">条件 A: 短系列</span>
  <span class="cg-body">系列長 512。標準的な NLU タスク。</span>
</div>

<div class="cg-card">
  <span class="cg-title">条件 B: 中系列</span>
  <span class="cg-body">系列長 2048。文書分類・要約タスク。</span>
</div>

<div class="cg-card">
  <span class="cg-title">条件 C: 長系列</span>
  <span class="cg-body">系列長 8192。長距離依存タスク。</span>
</div>

<div class="cg-card">
  <span class="cg-title">条件 D: 超長系列</span>
  <span class="cg-body">系列長 16384。本研究の主要ターゲット。</span>
</div>

</div>

---

<!-- _class: split-text -->

# [Split-Text] 分割テキストテンプレート

<div class="sp-left">
  <span class="sp-label">仮説</span>
  <span class="sp-body">動的スパースマスクを導入することで、計算量を削減しつつ精度を維持できる。</span>
</div>

<div class="sp-right">
  <span class="sp-label">結果</span>
  <span class="sp-body">3つのベンチマークで精度を維持しつつ、計算量を平均 60% 削減。</span>
</div>

---

<!-- _class: code -->

# [Code] コードテンプレート

<div class="cd-code">

```python
class SparseAttention(nn.Module):
    def __init__(self, d_model, n_heads):
        super().__init__()
        self.mask_gen = MaskGenerator(d_model)
        self.attn = nn.MultiheadAttention(
            d_model, n_heads
        )

    def forward(self, x):
        mask = self.mask_gen(x)
        return self.attn(x, x, x,
                         attn_mask=mask)
```

</div>

<div class="cd-desc">MaskGenerator がスパースマスクを動的生成し、標準の MultiheadAttention に渡す。</div>

---

<!-- _class: multi-result -->

# [Multi-Result] 複数結果テンプレート

<div class="mr-container">

<div class="mr-item">
  <span class="mr-metric">Perplexity</span>
  <span class="mr-value">17.9</span>
  <span class="mr-desc">Full Attention と同等の言語モデリング性能</span>
</div>

<div class="mr-item">
  <span class="mr-metric">Speedup</span>
  <span class="mr-value">2.4x</span>
  <span class="mr-desc">系列長 8192 での推論速度</span>
</div>

<div class="mr-item">
  <span class="mr-metric">Memory</span>
  <span class="mr-value">-40%</span>
  <span class="mr-desc">ピークメモリ使用量の削減率</span>
</div>

</div>

---

<!-- _class: takeaway -->

# [Takeaway] キーメッセージテンプレート

<div class="ta-main">動的スパース注意機構は、精度を犠牲にせず計算効率を大幅に改善できる</div>

<div class="ta-points">
<li>3つのベンチマークで SOTA と同等の精度を達成</li>
<li>計算量を平均 60% 削減し、長系列タスクへの適用を実現</li>
<li>既存の効率化手法と直交的に併用可能</li>
</div>

---

<!-- _class: profile -->

# [Profile] プロフィールテンプレート

<div class="pf-container">

<div class="pf-name">山田 太郎</div>

<div class="pf-affiliation">〇〇大学 工学研究科 情報工学専攻</div>

<div class="pf-bio">
<li>研究テーマ: 効率的な注意機構の設計と理論解析</li>
<li>所属: 自然言語処理研究室</li>
<li>学会活動: ACL 2025、NeurIPS 2024 Workshop</li>
<li>連絡先: yamada@example.ac.jp</li>
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
