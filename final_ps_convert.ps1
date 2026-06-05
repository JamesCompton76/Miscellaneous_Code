# powershell -ExecutionPolicy Bypass -File "H:\test\final_ps_convert.ps1"

# 1. Start the timer
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$word = New-Object -ComObject Word.Application
$word.Visible = $false

$folderPath = "H:\test"
$outputFolderPath = "H:\test\Convert"

# 2. Ensure the output folder exists
if (!(Test-Path -Path $outputFolderPath)) {
    New-Item -ItemType Directory -Path $outputFolderPath | Out-Null
}

# 3. Sees what files are in the folder
$files = Get-ChildItem -Path $folderPath | Where-Object { $_.Extension -in '.doc', '.docx' }

# 4. Try/Catch/Finally block to guarantee cleanup
try {
    foreach ($file in $files) {
        Write-Host "Converting: $($file.Name)..."
        
        $doc = $word.Documents.Open($file.FullName)
        
        $pdfFilename = [System.IO.Path]::ChangeExtension($file.Name, ".pdf")
        $outputPath = Join-Path -Path $outputFolderPath -ChildPath $pdfFilename
        
        $doc.SaveAs([String] $outputPath, [ref] 17) 
        
 # 5. Force close without saving changes to the original Word doc
        $doc.Close([ref] 0) 
    }
}
catch {
    Write-Warning "Script stopped due to an error: $_"
}
finally {
# 6. Ensures word closes regardless of success or failure
    if ($word) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        Remove-Variable word -ErrorAction SilentlyContinue
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
  
#7. Stop timer and display results
$stopwatch.Stop()
    Write-Host "----------------------------------------" -ForegroundColor Cyan
    Write-Host "Word application closed and memory released."
    Write-Host "Total execution time: $($stopwatch.Elapsed.Hours)h $($stopwatch.Elapsed.Minutes)m $($stopwatch.Elapsed.Seconds)s" -ForegroundColor Green
    Write-Host "----------------------------------------" -ForegroundColor Cyan
}