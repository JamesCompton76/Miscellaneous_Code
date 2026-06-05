# Miscellaneous_Code
Powershell, VBA etc.

1.  final_ps_convert.ps1 - takes all the .doc and .docx files in a specified folder and converts them to pdf.  Added in code for timer, total time, files as they go etc.  Note, whether this is done in Powershell or Python it does rely on the Word .com object existing on the local pc, though this sometimes exists if there's a trial version etc. in the background.  Can be rewritten to use Libre if necessary with a slight change, according to Gemini that's along the lines of the below, though I haven't tested that.

Define the path to LibreOffice

$soffice = "C:\Program Files\LibreOffice\program\soffice.exe"

Run the headless converter

& $soffice --headless --convert-to pdf --outdir "C:\Your\Output\Folder" "C:\Your\Input\Folder\document.docx"
