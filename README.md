# PowerShell_Utilities
PowerShell Scripts (Automation & Conversion)

## Highlighted Utilities

1. **[final_ps_convert.ps1](./final_ps_convert.ps1)** - Headless COM-object automation script with strict memory management for batch converting Microsoft Word documents to PDF.
   <details>
   <summary><i>Click to expand architectural breakdown</i></summary>

   An automation pipeline that orchestrates Microsoft Word via COM objects to perform headless batch conversions of `.doc` and `.docx` files to `.pdf`. It implements rigorous memory management and error handling to prevent resource leaks and orphaned background processes during high-volume conversions.

   * **Headless COM Orchestration:** Instantiates the `Word.Application` COM object in the background (`$word.Visible = $false`) to process document conversions silently without disrupting the desktop environment.
   * **Guaranteed Resource Cleanup:** Utilizes a strict `try/catch/finally` control flow alongside explicit Garbage Collection (`[System.GC]::Collect()`) and COM object release (`[System.Runtime.Interopservices.Marshal]::ReleaseComObject`). This ensures the Word background process terminates safely and memory is freed, regardless of individual file errors.
   * **Execution Profiling:** Implements `System.Diagnostics.Stopwatch` to track and output precise telemetry on the batch run's total execution time.
   * **Execution Note:** To run this script and temporarily bypass local system execution constraints, use the following command:
     `powershell -ExecutionPolicy Bypass -File "H:\test\final_ps_convert.ps1"`
   </details>
