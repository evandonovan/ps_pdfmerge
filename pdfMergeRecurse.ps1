# Script for command line or automated calling of PDF merges
# Call using:
# powershell.exe -ExecutionPolicy Bypass -Command "& 'C:\Users\USERNAME\Software\pdfMerge\pdfMergeRecurse.ps1'  -srcDir C:\pdfSrc -destDir C:\pdfDest"

# http://9to5it.com/bypass-the-powershell-execution-policy/

param (
 [string] $srcDir,
 [string] $destDir
)

# load the module (adjust path as needed)
Import-Module -Name "C:\Users\USERNAME\Software\pdfMerge\cvu.pdfMergeRecurse.psm1"
Get-Module MergePDFRecurse

# call using parameters (which must be set from command line, or batch job)
MergePDFRecurse $srcDir $destDir
