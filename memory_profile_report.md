# python-docx — W11-B memory profile report

Recorded via `resource.getrusage(RUSAGE_SELF).ru_maxrss` (high-water RSS, KiB) plus `tracemalloc` peak (Python heap, bytes). Each scale runs in a fresh subprocess so `ru_maxrss` reflects that fixture alone.

The `Library-attributable` column is `peak RSS − pre-import baseline`: it excludes the cold Python interpreter (~30–40 MiB on Linux) which is not the library's responsibility. This is what we compare against the on-disk fixture size when flagging outliers (>5x).

| Scale | Paragraphs | Fixture (KiB) | Baseline RSS (MiB) | Peak RSS (MiB) | Lib-attributable (MiB) | Peak Py heap (MiB) | Lib RSS / fixture |
|---|---|---|---|---|---|---|---|
| 100p | 100 | 20.9 | 12.8 | 36.7 | 23.9 | 0.9 | 1171.8x  FLAG |
| 1k | 1,000 | 39.5 | 12.6 | 42.0 | 29.5 | 2.0 | 763.8x  FLAG |
| 5k | 5,000 | 122.0 | 12.6 | 68.0 | 55.5 | 8.0 | 465.6x  FLAG |
| 10k | 10,000 | 225.0 | 12.6 | 102.8 | 90.2 | 10.6 | 410.7x  FLAG |

## Per-checkpoint RSS deltas (MiB above pre-import baseline)

| Scale | pre-load | post-load | post-manip | post-save |
|---|---|---|---|---|
| 100p | 17.4 | 23.2 | 23.2 | 23.9 |
| 1k | 17.4 | 28.2 | 28.2 | 29.5 |
| 5k | 16.9 | 50.0 | 50.7 | 55.5 |
| 10k | 17.3 | 77.7 | 80.2 | 90.2 |

## Outliers (library-attributable RSS > 5x fixture size)

- **100p** (100 paragraphs): library-attributable RSS 23.9 MiB vs fixture 20.9 KiB on disk (1171.8x)
- **1k** (1,000 paragraphs): library-attributable RSS 29.5 MiB vs fixture 39.5 KiB on disk (763.8x)
- **5k** (5,000 paragraphs): library-attributable RSS 55.5 MiB vs fixture 122.0 KiB on disk (465.6x)
- **10k** (10,000 paragraphs): library-attributable RSS 90.2 MiB vs fixture 225.0 KiB on disk (410.7x)
