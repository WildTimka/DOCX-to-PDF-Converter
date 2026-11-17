# DOCX-to-PDF Converter v1.1.0

- Adaptive Microsoft Word worker pool, automatically limited by document count, logical CPUs, available memory and measured startup cost.
- Sequential COM initialization overlaps ongoing conversion; a dynamic shared queue balances documents across isolated workers.
- Controlled startup and single-worker fallback, per-batch reuse, pinned owned-process handles and cancellation cleanup.
- Improved document-proxy closing after RPC_E_DISCONNECTED, retry handling and staged PDF publication/rollback.
- Native compact Windows GUI accumulated since the v1.0.0 console release: folder selection, visible progress, Cancel and result display. The validated GUI layout is unchanged in this release preparation.
- Print-quality PDF export and existing PDF overwrite behavior retained; original documents remain unchanged and the previous default printer is restored on handled completion/cancellation.

## Historical performance evidence

Three-run medians per configuration from 48 matched end-to-end research trials (2,004 successful conversions). The baseline is a corrected fixed one/two-worker engine sharing reliability fixes with the adaptive candidate, not a new test of the unmodified v1.0.0 executable. These are historical measurements, separate from the focused release regression checks.

| Documents | Baseline | Adaptive | Speedup |
|---:|---:|---:|---:|
| 5 | 25.585 s | 16.348 s | 1.57× |
| 10 | 46.077 s | 22.933 s | 2.01× |
| 20 | 50.515 s | 35.367 s | 1.43× |
| 40 | 94.802 s | 59.395 s | 1.60× |
| 80 | 180.927 s | 97.343 s | 1.86× |
| 176 | 386.564 s | 196.500 s | 1.97× |

Measured on Intel Core i7-6700K, 4 cores / 8 logical processors, approximately 32 GiB RAM, Windows 10 Pro, Microsoft Word x86 16.0.20326.20232. Fresh Word processes, warm OS caches, isolated E: research staging. Performance on other systems and workloads may differ; the application is not universally twice as fast.

Target workers: 1→1; 2→2; 3–5→3; 6–29→4; 30–59→5; 60–79→6; 80+→7. Resource and startup guards may lower these targets. Exact break-even points between sampled sizes and other hardware are not validated. Periodic restart has no demonstrated benefit.

## Validation and known limitations

Historical automated/core tests covered successful conversion, retries, failed-worker fallback, ownership, active/startup cancellation, overwrite, cache cleanup and printer restoration. Historical PDF QA validated 13,925 PDFs / 225,211 pages and 15 matching representative raster pages. These counts are not new release-test results.

The owner additionally reports successful manual GUI conversion, stable behavior and correctly terminated Word processes. This is user-reported validation. The standalone hidden pre-existing Word test preserved its process/document twice, but its fixture's own Quit failed; exact-identity cleanup was subsequently required. Those overall fixture tests remain recorded as failed and are not presented as fully passed.

Desktop Microsoft Word is required. Printer changes are temporarily system-wide; avoid concurrent batches and forced termination. Cold caches, all Word versions, unusual documents and aging beyond 176 files are not exhaustively validated. A COM activation hang before ownership identification remains an unproven failure mode; unidentified/unrelated Word processes must never be killed.
