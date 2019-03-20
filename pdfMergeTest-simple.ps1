# merges files from src to dst directory on main drive - no recursion across multiple directories
# run using powershell.exe -ExecutionPolicy Bypass -Command "& 'C:\Users\USERNAME\Scripts\pdfMerge\pdfMergeTest-simple.ps1'
$PdfMergeModule = "cO.Pdfsharp"
$PdfMergeModulePath = "C:\Users\USERNAME\Scripts\pdfMerge\cO.Pdfsharp.psm1"
$PdfSharpPath = "C:\Users\USERNAME\Scripts\pdfMerge\PdfSharp-gdi.dll"
$srcDir = "C:\src"
$dstDir = "C:\dst"
Import-Module -Name $PdfMergeModulePath
Get-Module $PdfMergeModule
$files = Get-ChildItem $srcDir -Filter "*.pdf"
MergePdf -Files $files -DestinationFile "C:\dst\merged.pdf" -PdfSharpPath $PdfSharpPath -Force
