<#
This module provides functionality to merge multiple PDF files from under subdirectories, and save in a directory of choice. 
# Script by Evan Donovan, created on behalf of City Vision University (www.cityvision.edu)
See README.md for details.
It was tested with PdfSharp-gdi.dll in version 1.5.4 and Powershell 5 and is provided "as is" under the GPL.
#>
<#
.SYNOPSIS
Merges multiple PDF files under child directories into one PDF per directory, using that directory's name
and saves in a destination directory.
.PARAMETER srcDir
Directory under which the child directories with PDFs are.
.EXAMPLE
$srcDir = "C:\Users\USERNAME\pdfSrc"
$destDir = "C:\Users\USERNAME\pdfDest"
MergePdfRecurse $srcDir $destDir
#>
function MergePDFRecurse {
# TODO: validate whether these are both existant directories
Param(
  [Parameter(Mandatory=$True,Position=1)]
   [string]$srcDir,
	
   [Parameter(Mandatory=$True,Position=2)]
   [string]$destDir
)
 
    # define name of PdfMerge module
    $PdfMergeModule = "cO.Pdfsharp"

    # 0) Set up the script - must change these before running on your system

    # ** change this to where you want the pdfs
    $PdfMergeModulePath = "C:\Users\USERNAME\Software\pdfMerge\cO.Pdfsharp.psm1"
    $PdfSharpPath = "C:\Users\USERNAME\Software\pdfMerge\PdfSharp-gdi.dll"

    # 1) load the needed module
    Import-Module -Name $PdfMergeModulePath
    Get-Module $PdfMergeModule

    # 2) collect all the pdfs in each sub-directory for processing:
    # only one level of recursion
    # https://blogs.technet.microsoft.com/heyscriptingguy/2014/02/03/list-files-in-folders-and-subfolders-with-powershell/
    $srcDirs = Get-ChildItem -Path $srcDir -Directory
    foreach($childDir in $srcDirs) {
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
}
