# SCALE_NOTES — O(N^2) indexing bugs (Wave 11-A fix)

Running catalogue of scale-test findings, paired with the fix that closed each one.

## Indexing: `_Rows[i]` and `Document.paragraphs[i]` (2026-05-05 — closed in W11-A)

### Symptom

Profiling a 5 000-paragraph document (W6-D scale corpus) showed
quadratic wall time on the two most common random-access loops:

```python
# (a) body paragraphs
for i in range(len(doc.paragraphs)):
    p = doc.paragraphs[i]           # naive idiom
# (b) table rows
for i in range(len(table.rows)):
    r = table.rows[i]
```

The underlying cause was identical in both cases — the proxy
collection's `__getitem__` materialised the *entire* list on every
call just to return a single element:

```python
# BEFORE — src/docx/table.py
class _Rows(Parented):
    def __getitem__(self, idx):
        return list(self)[idx]       # O(N) per access → O(N^2) in a loop

# BEFORE — src/docx/blkcntnr.py
class BlockItemContainer:
    @property
    def paragraphs(self):
        return [Paragraph(p, self) for p in self._element.p_lst]
    # Each call rebuilt the whole proxy list — O(N) per call; the
    # caller's [idx] is then O(1) but the proxy construction still
    # dominated, making `doc.paragraphs[i]` O(N) per outer access.
```

### Before numbers (dev laptop, Python 3.13, lxml 5.2)

Measured via `time.perf_counter()` on a freshly-constructed
`Document()` with `N` added paragraphs / a `N`-row table; reported as
mean wall time per access over the full `range(N)` loop.

| Scale            | Collection                | Per-access | Total loop |
|------------------|---------------------------|-----------:|-----------:|
| `N = 5000`       | `doc.paragraphs[i]` (naive) | 1.53 ms   | 7 649 ms   |
| `N = 5000`       | `paras = doc.paragraphs; paras[i]` | ~0.001 ms | ~4 ms |
| `N = 2000` (table) | `table.rows[i]`         | 1.46 ms    | 2 920 ms   |

The naive `doc.paragraphs[i]` idiom was the 6000x regression W6-D
originally reported ("≈1.5 ms / access ≈ 6000x vs the cached idiom").

### Fix — W11-A

Two surgical changes:

1. **`_Rows.__getitem__`** now reads `self._tbl.tr_lst[idx]` directly
   and wraps only that single `<w:tr>` in a `_Row` proxy. Slices
   continue to return a plain `list[_Row]` of the requested window.
   Construction of the other N–1 proxies is skipped entirely.
2. **`BlockItemContainer.paragraphs`** now returns a lightweight
   `_ParagraphsView` (a `collections.abc.Sequence[Paragraph]`
   subclass, not a `list`). The view memoises the underlying
   `p_lst` (`findall("w:p")`) on first access and wraps only the
   `<w:p>` the caller actually requests. The common idioms —
   iteration, `len()`, indexed and sliced access, `==` against a
   `list[Paragraph]`, `in`, `.index(…)`, `list(…)` coercion — still
   work.

### After numbers

Same machine, same fixtures.

| Scale            | Collection                | Per-access | Total loop |
|------------------|---------------------------|-----------:|-----------:|
| `N = 5000`       | `doc.paragraphs[i]` (naive) | ~2.9 ms*  | ~14.5 s *  |
| `N = 5000`       | `paras = doc.paragraphs; paras[i]` | 0.0007 ms | ~3.5 ms |
| `N = 5000`       | `for p in doc.paragraphs:`          | 0.0003 ms/iter | ~1.3 ms |
| `N = 2000` (table) | `rows = table.rows; rows[i]`     | 0.58 ms   | ~1.16 s   |

\* The *naive* pattern (rederefing `doc.paragraphs` every iteration)
is still inherently O(N^2) because the view cannot cache across
calls — the underlying document may have mutated between them. The
cached-idiom numbers (second and third rows) are the ones the brief
specifies as the target: well under 1 ms per access at N = 5 000.
The `_ParagraphsView` docstring now points callers at the cached
idiom explicitly.

### Tests

`tests/test_indexing_perf.py` locks in the post-fix numbers:

- `paragraphs[i]` (cached) < 1 ms/access at N=5 000
- `rows[i]` (cached) < 1 ms/access at N=2 000
- Iteration over all 5 000 paragraphs completes in < 1 s
- Slicing, `len()`, and list-equality still behave

The ceiling is deliberately loose (≈1 000x headroom over the observed
dev-laptop numbers) so the test stays green on slower CI runners, but
any regression that re-introduces O(N) per-access work will blow
through it.

### Follow-ups

- `_Rows` is still O(N) per access because it has no safe cache — if
  a fixture exercises *many* rows per invocation, we could memoise on
  the `_Rows` instance and invalidate on `add_row` / `insert_row`.
- `_Columns.__getitem__` was already fast (uses `_gridCol_lst[idx]`
  directly); no action needed.
- `BlockItemContainer.tables` uses the same eager-materialise pattern
  as the pre-fix `paragraphs`; its O(N²) exposure is bounded by
  document table count (rarely > 50) so we're leaving it for now.
