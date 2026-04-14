---
marp: true
theme: academic
paginate: true
---

<!-- _class: code -->

# コード例

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

<div class="cd-desc">MaskGenerator がスパースマスクを動的生成し、標準の MultiheadAttention に渡す。既存コードへの変更は最小限。</div>
