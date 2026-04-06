---
marp: true
theme: academic
paginate: true
---

<!-- _class: table-slide -->

# 表解説スライド

## 各手法の性能比較

| 手法 | Accuracy | F1 Score | 推論時間 (ms) | パラメータ数 |
|------|:--------:|:--------:|:------------:|:----------:|
| Baseline | 89.2% | 0.881 | 12 | 11M |
| Method A | 92.1% | 0.915 | 18 | 25M |
| Method B | 93.4% | 0.928 | 45 | 110M |
| **Ours** | **96.7%** | **0.962** | **15** | **18M** |

<div class="box-accent">

**提案手法の優位性**: 最高精度を達成しつつ、推論時間・パラメータ数ともに効率的。

</div>

<div class="footnote">すべての実験は同一のハードウェア環境 (NVIDIA A100) で実施</div>
