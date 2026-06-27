# Miscellaneous_Code
A collection of utility scripts and automation tools written in PowerShell, Batch (.bat), VBScript, and other languages.

## PowerShell Scripts

1. **[final_ps_convert.ps1](./final_ps_convert.ps1)** - Headless COM-object automation script with strict memory management for batch converting Microsoft Word documents to PDF.
   
   <details>
   <summary><i>Click to expand architectural breakdown</i></summary>

   An automation pipeline that orchestrates Microsoft Word via COM objects to perform headless batch conversions of `.doc` and `.docx` files to `.pdf`. It implements rigorous memory management and error handling to prevent resource leaks and orphaned background processes during high-volume conversions.

   * **Headless COM Orchestration:** Instantiates the `Word.Application` COM object in the background (`$word.Visible = $false`) to process document conversions silently without disrupting the desktop environment.
   * **Guaranteed Resource Cleanup:** Utilizes a strict `try/catch/finally` control flow alongside explicit Garbage Collection (`[System.GC]::Collect()`) and COM object release (`[System.Runtime.Interopservices.Marshal]::ReleaseComObject`). This ensures the Word background process terminates safely and memory is freed, regardless of individual file errors.
   * **Execution Profiling:** Implements `System.Diagnostics.Stopwatch` to track and output precise telemetry on the batch run's total execution time.
   * **Execution Note:** To run this script and temporarily bypass local system execution constraints, use the following command:
     `powershell -ExecutionPolicy Bypass -File "H:\test\final_ps_convert.ps1"`
   * **Further Note:** This relies on the local pc having the Word COM object available, can be updated to use LibreOffice if necessary but will require some changes.
   </details>

## Stream Deck & Batch Scripts

1. **[LaunchDev.bat](./LaunchDev.bat)** - One-click automation script to launch WSL2 Ubuntu, activate a Python virtual environment, and open VS Code via Elgato Stream Deck.
   
   <details>
   <summary><i>Click to expand architectural breakdown</i></summary>

   A batch script and configuration workflow designed to pass multiple execution steps into a Linux subsystem from a native Windows shortcut, maintaining an interactive terminal session while preventing Conda from automatically overwriting the target environment.

   * **WSL2 Orchestration:** Utilizes a temporary initialization script (`/tmp/wsl_init.sh`) to dynamically load the bash profile (`~/.bashrc`), activate the virtual environment (`.venv`), and launch VS Code before replacing the automated shell with a fully interactive terminal session (`--rcfile /tmp/wsl_init.sh -i`).
   * **Conda Environment Management:** Requires disabling Conda's default terminal hijacking by running `conda config --set auto_activate_base false` once in a standard Ubuntu terminal. This ensures VS Code's integrated terminal correctly inherits the intended `.venv` instead of reverting to `(base)`.
   * **VS Code Integration:** Once opened via the script, use `Ctrl + Shift + P` -> **Python: Select Interpreter** to map the editor to the `.venv` executable.
   * **Stream Deck Setup:** Implemented by mapping the **System -> Open** action to the `.bat` file on your Stream Deck.
   * **Execution Note:** The script code for the `.bat` file is:
     ```bat
     "C:\Program Files\WSL\wsl.exe" --distribution-id {b3c166ea-97b5-46d8-ab70-769e8adb311f} --cd ~ -e bash -c "echo 'source ~/.bashrc; source .venv/bin/activate; code .' > /tmp/wsl_init.sh && bash --rcfile /tmp/wsl_init.sh -i"
     ```
   * **Further Note (VS Code Warnings):** You may occasionally see an orange warning (`!`) stating extensions want to relaunch the terminal. **You can safely ignore this.** It occurs because the automated terminal opens slightly faster than the Python extension loads in the background. Do not click "Relaunch terminal", as doing so forces a hard reset that may drop the `.venv`. If you accidentally drop the environment, simply close the terminal (Trash Can icon) and open a new one (`` Ctrl + Shift + ` ``) to properly reload into `.venv`.
   </details>
