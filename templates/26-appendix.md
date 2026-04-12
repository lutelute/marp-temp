---
marp: true
theme: academic
paginate: true
---

<!-- _class: appendix -->

# ハイパーパラメータ一覧

<span class="appendix-label">Appendix A</span>

| パラメータ | 値 | 備考 |
|---|---|---|
| 隠れ次元 | 768 | BERT-base と同一 |
| ヘッド数 | 12 | |
| 層数 | 12 | |
| 学習率 | 3e-4 | cosine schedule |
| バッチサイズ | 64 | |
| Dropout | 0.1 | |
| スパース率 $k$ | $\sqrt{n}$ | 動的調整あり |
