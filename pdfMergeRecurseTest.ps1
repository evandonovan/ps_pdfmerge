# script to merge all the pdfs by name
# uses CoolOrange's mergepdf powershell module
# which uses a compiled DLL of the open-source .NET PDFSharp extension
# first install their script from here: 
# https://support.coolorange.com/support/solutions/articles/22000211000-how-to-merge-multiple-pdf-with-pdfsharp-and-powershell

# Script by Evan Donovan, created on behalf of City Vision University (www.cityvision.edu)
# Code licensed under GPL

# See README.md for usage

$PdfMergeModule = "cO.Pdfsharp"

# 0) Set up the script - must change these before running on your system

# ** change this to where you want the pdfs
$PdfMergeModulePath = "C:\Users\USERNAME\SOFTWARE\pdfMerge\cO.Pdfsharp.psm1"
$PdfSharpPath = "C:\Users\USERNAME\SOFTWARE\pdfMerge\PdfSharp-gdi.dll"

# ** change to where the source pdfs live
$srcDir = "C:\Users\USERNAME\pdfSrc"

# ** change to where you want the final pdfs
$destDir = "C:\Users\USERNAME\pdfDest"

# 1) load the needed module
Import-Module -Name $PdfMergeModulePath
Get-Module $PdfMergeModule

# 2) collect all the pdfs in each sub-directory for processing:
# only one level of recursion
# https://blogs.technet.microsoft.com/heyscriptingguy/2014/02/03/list-files-in-folders-and-subfolders-with-powershell/
$srcDirs = Get-ChildItem -Path $srcDir -Directory
foreach($childDir in $srcDirs)
{
  $childDirFullPath = $srcDir + "/" + $childDir.Name
  # sort the files using a regex so they sorted the same as in Windows Explorer
  # https://stackoverflow.com/questions/5427506/how-to-sort-by-file-name-the-same-way-windows-explorer-does
  $files = Get-ChildItem $childDirFullPath -Filter "*.pdf" | Sort-Object { [regex]::Replace($_.Name, '\d+', { $args[0].Value.PadLeft(20) }) }

  $destFilename = $childDir.Name + ".pdf"
  $destFile = $destDir + "\" + $destFilename

  # 3) merges the pdfs using the PDFSharp extension
  # removed the -RemoveSourceFiles parameter since not needed
  # -Force is used to override read-only
  MergePdf -Files $files  -DestinationFile $destFile -PdfSharpPath $PdfSharpPath -Force
}
