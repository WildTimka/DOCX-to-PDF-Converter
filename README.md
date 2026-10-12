# DOCX to PDF Converter

Batch convert Microsoft Word DOCX documents into print quality PDFs on Windows x64. The application automates Microsoft Word through COM. **Microsoft Word is required for conversion.** Python is not required.

## Download

Download the validated Windows executable from the [DOCX to PDF Converter releases](https://github.com/WildTimka/DOCX-to-PDF-Converter/releases/latest). Place `DOCX_to_PDF_FINAL_x64.exe` in the folder containing the DOCX files, then run it. PDFs are written beside their source documents.

## Features

- Finds DOCX files directly alongside the executable; it does not search subfolders and skips Word lock files whose names start with `~$`.
- Processes up to 10 documents sequentially with one Word worker. With more than 10 documents, it starts two Word workers and distributes jobs from a shared queue as workers become available.
- Allows up to 30 seconds for Word workers to start. If the second worker cannot start, processing continues with one worker.
- Restarts a worker and retries a failed document once. Remaining queued work can continue on another live worker; individual failures are reported.
- Uses Microsoft Word's fixed format PDF export with print quality settings.
- Prints progress, per document status, totals, and total elapsed time in the console.
- Runs as a standalone Windows x64 executable without Python or third party Go packages.

## Requirements

- Windows 10 or later, 64 bit.
- Microsoft Word desktop installed, activated, and available through COM automation.
- Write permission in the folder containing the documents.
- For documents on another drive, free space and write permission in a temporary directory on drive C:.
- Microsoft Print to PDF when the current default printer cannot be used for PDF export.

## Quick start

1. Install and open Microsoft Word once so initial setup dialogs are complete.
2. Download `DOCX_to_PDF_FINAL_x64.exe` from the [latest release](https://github.com/WildTimka/DOCX-to-PDF-Converter/releases/latest).
3. Copy the executable next to the DOCX documents to convert.
4. Run the executable and leave its console open until it reports completion.
5. Review the created PDF files and any per document errors in the console.

The application uses the executable's folder as its input folder. For example, if `C:\\Invoices\\DOCX_to_PDF_FINAL_x64.exe` is beside `January.docx`, it writes `January.pdf` in `C:\\Invoices\\`.

## Parallel processing

For 10 or fewer DOCX files, one Word worker processes documents in sequence. For 11 or more, the program attempts to run two separate Word COM workers. They draw jobs from a shared queue, so a faster worker can take more documents. When the second Word instance fails to start, the first worker handles the batch alone. A failed conversion is retried once after restarting its worker; a persistent failure is recorded while remaining jobs continue where possible.

## Printer management

The program reads the current default printer and selects Microsoft Print to PDF if its printer check determines a switch is needed. It restores the saved default printer after processing, including ordinary conversion errors. Printer selection is system wide while the application runs, so wait for completion and avoid forcibly terminating the program.

## Temporary file handling

When the source folder is on C:, documents are processed in place without an extra cache copy. When it is on another drive, the program stages DOCX files in a temporary directory on C:, processes those copies, then publishes the PDFs beside the originals. The temporary cache is removed after processing.

## Build from source

Install Go 1.23 or later on Windows, then run this from the project root:

```powershell
.\build\build.ps1
```

The script sets `GOOS=windows`, `GOARCH=amd64`, and `CGO_ENABLED=0`. It builds from `src/`, where `word_worker.ps1` is embedded into the Go executable through `//go:embed`. The output is `DOCX_to_PDF_REBUILT_x64.exe` in the project root. Word is not needed to compile, but is required when running the converter.

Builds from reconstructed source are for development and evaluation. Test them on Windows with Word and a representative batch before distributing them. The build script exits with an error if `go build` fails.

## Troubleshooting

- **No files found:** Make sure the executable is in the same folder as the `.docx` files. Nested folders are not scanned; lock files beginning with `~$` are ignored.
- **Word does not start:** Install/activate desktop Microsoft Word, launch it once interactively, and confirm Windows can create a Word COM automation instance.
- **PDF export fails:** Try opening and exporting the document directly in Word. Check that a usable printer is installed and that Microsoft Print to PDF is enabled if needed.
- **Files on another drive fail:** Check free space on C: and write access to the user's temporary directory.
- **A document reports an error:** Close it in Word and rerun it individually. Complex or damaged documents may fail or exceed the processing timeout.
- **Printer was not restored:** If the application was terminated during processing, select the prior printer again in Windows printer settings.

## Known limitations

- Windows x64 only; Microsoft Word desktop is mandatory.
- Only DOCX files in the executable's immediate directory are processed.
- Word dialogs, damaged documents, unusual COM registration, or long running documents can interrupt conversion.
- Printer changes are system wide until processing finishes; forced termination may bypass normal restoration.
- Output PDFs with matching names may be overwritten.
- Performance varies with Word, the documents, storage, and the computer.

## Source provenance and validation

The original Go source used for the validated executable was not preserved. `src/main.go` was reconstructed from the released executable's embedded PowerShell worker and observed behavior. `src/word_worker.ps1` is the extracted worker. The source is available for continued development but is not represented as the exact historical source or as behaviorally equivalent to the release.

The supplied `DOCX_to_PDF_FINAL_x64.exe` is the known validated version. The project owner reports it converted 177 DOCX documents with zero errors in about 3 minutes 28 seconds on their Windows computer. This result applies to that executable. Reconstructed source builds have not been functionally validated with Windows and Word and must not be treated as stable releases until tested. No screen optimized PDF mode or development debug output is included.
