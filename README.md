# DOCX-to-PDF Converter v1.1.0

Batch convert DOCX documents into print-quality PDFs with a compact native Windows GUI. Microsoft Word desktop is required. The standalone Windows x64 executable needs neither Go nor Python installed.

## Requirements and use

Windows 10/11 x64, installed and activated Microsoft Word desktop (32-bit or 64-bit), a usable printer, write access beside the documents, and free temporary space on C:. Open Word once to finish first-run dialogs before conversion.

1. Download the Windows x64 executable from the GitHub releases page.
2. Run it, browse to a folder or enter its path, then select Convert.
3. Follow the progress counter. PDFs are written beside the original documents; existing PDFs with matching names are overwritten after conversion completes.
4. Select Cancel, or close the window, to request cancellation and wait for cleanup. Cancellation before publication preserves existing PDFs and publishes no new PDFs from that batch.

Only DOCX files directly in the selected folder are included; subfolders and Word lock files beginning with ~$ are skipped. Original DOCX files are not modified. No concurrency settings are necessary.

## Adaptive Word workers

The engine starts converting as soon as the first isolated Word worker is ready. Additional Word COM instances initialize sequentially while ready workers draw documents from a shared queue. Workers are reused for the batch and replaced after conversion failures when appropriate; a failed document is retried once. Additional-worker startup failure leaves existing workers processing the queue.

| Documents | Target workers |
|---|---:|
| 1 | 1 |
| 2 | 2 |
| 3–5 | 3 |
| 6–29 | 4 |
| 30–59 | 5 |
| 60–79 | 6 |
| 80+ | 7 |

These are upper targets calibrated on one workload, not universal optima. Logical CPU count, available memory (2 GiB reserve and 512 MiB per worker), verified process ownership, and observed startup cost versus remaining work can reduce concurrency. Unknown resource capacity falls back to one worker. There are no periodic worker restarts.

Word process handles are pinned after ownership verification; unrelated user Word instances must never be terminated. Each worker owns its own STA COM object. The engine stages PDFs before publication, removes temporary conversion files, and restores the saved default printer after completion, cancellation or handled failure. Printer selection is temporarily system-wide: run one batch at a time and avoid forcibly terminating the application. Non-C: input files are staged in a writable C: temporary directory.

## Build and test

Install Go 1.23 or later on Windows. The release was built with Go 1.26.9 windows/amd64. From this source root:

```powershell
.\build\test.ps1
.\build\build.ps1
```

Output: dist/DOCX_to_PDF_Converter_v1.1.0_Windows_x64.exe. Both scripts accept -Go with a compiler path; build also accepts -Output. The build embeds src/word_worker.ps1, disables CGO, trims paths and uses the Windows GUI subsystem (no console). No third-party Go modules or application runtime dependencies are added. Buildvcs is disabled for byte-for-byte reproduction from the source ZIP as well as Git; the release manifest associates the artifact with its exact commit and hashes. Reproduction requires the same compiler version and build flags.

The automated suite does not require Word. Real conversion validation requires Windows and Word; see TESTING.md. VERSION and the linker version value must agree when preparing another release.

## Troubleshooting and limitations

- No documents: select the folder containing DOCX files, not its parent.
- Word startup failure: complete Word activation/first-run prompts and verify documents open normally. Additional-worker failure falls back to the ready workers.
- Export failure: open/export the document in Word; damaged, protected, interactive or unusually complex documents can fail. Worker startup timeout is 30 seconds; document timeout is 5 minutes.
- Temporary-folder failure: check free space and write permissions on C:.
- Printer changed after forced termination: reselect the prior default printer in Windows. Normal cancellation waits for restoration; a forced process/OS termination cannot guarantee cleanup.
- Performance varies with documents, Word build, storage and background load. Seven is a practical tested ceiling, not a guarantee on every machine. A COM activation hang before ownership can be established remains a known limitation; unidentified processes are never forcibly terminated.

See RELEASE_NOTES_v1.1.0.md for historical measurements and validation limits. v1.0.0 remains preserved with its original assets; its reconstructed source is historical. This release uses the later GUI v5 source and validated adaptive engine, rather than claiming to reproduce the v1.0.0 binary.
