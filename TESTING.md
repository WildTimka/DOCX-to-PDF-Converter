# Release validation procedure

Run build/test.ps1 for all automated tests and go vet; then build/build.ps1. On Windows with activated Word, use isolated copies of representative documents for small, medium and 80+ document batches. Verify PDFs parse, page counts/mapping agree, originals remain unchanged, overwrite succeeds, progress advances, cancellation preserves existing PDFs, no owned workers or temporary caches survive and the original printer is restored. Check window-close cancellation and preservation of an unrelated manually opened Word document. Do not kill processes by name.

User-reported manual validation for v1.1.0: GUI conversion works reliably and Word processes close correctly. It is distinct from automated/core regression evidence. Final release checks and exact build/artifact hashes are recorded in the release validation report and manifest; historical research evidence is archived locally rather than bundling datasets or PDF collections here.

The source GUI is byte-identical to the validated GUI v5 layout. Production C: staging remains unchanged. E:-only research harness copies override the cache allocator and add observation events; their conversion results validate the same engine code but are not measurements of GUI rendering or production C: staging.
