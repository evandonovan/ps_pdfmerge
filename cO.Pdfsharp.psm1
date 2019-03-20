<#
This module provides functionality to merge multiple pdf files. It was tested with PdfSharp-gdi.dll in version 1.5.4 and PowerShell 5 and is provided "as is"

Evan Donovan has modified for CityVision.edu to add a footer to each document page
with the file name and page number (for the merged document as a whole), 
and also to add bookmarks for each document added in order to easily identify page numbers.
#>
function MergePdf {
<#
.SYNOPSIS
Merges multiple PDF files into a single multisheet PDF
.PARAMETER Files
A collection of PDF files
.PARAMETER DestinationFile
Full path of the merged PDF
.PARAMETER PdfSharpPath
Full path to the required PdfSharp library
.PARAMETER Force
Tries to force destination file and directory creation and deletion of source files, even when they are read-only
.PARAMETER
RemoveSourceFiles
Deletes the source files after PDF is merged
.EXAMPLE
$files = Get-ChildItem "C:\temp\PDF\Source" -Filter "*.pdf"
MergePdf -Files $files  -DestinationFile "C:\TEMP\PDF\Destination\test.pdf" -PdfSharpPath 'C:\ProgramData\coolOrange\powerJobs\Modules\PdfSharp-gdi.dll' -Force -RemoveSourceFiles
#>
param(
[Parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]
[ValidateScript({
    if( $_.Extension -ine ".pdf" ){ 
        throw "The file $($_.FullName) is not a pdf file."
    } 
    if(-not (Test-Path $_.FullName)) {
        throw "The file '$($_.FullName)' does not exist!"
    }
    $true
})]
[System.IO.FileInfo[]]$Files,
[Parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]
[System.IO.FileInfo]$DestinationFile,
[Parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]
$PdfSharpPath,
[switch]$Force,
[switch]$RemoveSourceFiles
)
    Write-Host ">> $($MyInvocation.MyCommand.Name) >>"

    if((Test-Path $PdfSharpPath) -eq $false) {
        throw "Could not find PdfSharp assembly at $($PdfSharpPath)"
    }
    Add-Type -LiteralPath $PdfSharpPath

    if((Test-Path $DestinationFile.FullName) -and $DestinationFile.IsReadOnly -and -not $Force) {
        throw "Destination file '$($DestinationFile.FullName)' is read only"
    }

    [System.IO.DirectoryInfo]$DestinationDirectory = $DestinationFile | Split-Path -Parent
    if(-not (Test-Path $DestinationDirectory)) {
        try {
            $DestinationDirectory = New-Item -Path $DestinationDirectory.FullName -ItemType Directory -Force:$Force
        } catch {
            throw "Error in $($MyInvocation.MyCommand.Name). Could not create directory '$($Path)'. $Error[0]"
        }
    }

    # initialize the output document into which the pdf will be saved
	# http://www.pdfsharp.net/wiki/ConcatenateDocuments-sample.ashx
    $pdf = New-Object PdfSharp.Pdf.PdfDocument
	
	# initialize variables for footer labelling document
	# http://www.pdfsharp.net/wiki/CombineDocuments-sample.ashx	
	# using the Powershell syntax documented at:
	# https://merill.net/2013/06/creating-pdf-files-dynamically-with-powershell/
	$font = New-Object PdfSharp.Drawing.XFont('Verdana', 9, [PdfSharp.Drawing.XFontStyle]::Bold)
    $box = New-Object PdfSharp.Drawing.XRect
	
    #Write-Host "Creating new PDF"
	
	# page number starts at 1
	$page_number = 1
	
	
	
	# iterate over the files being merged
    foreach ($file in $Files) {
	    # open the file (need to open for Import in order to do addPage)
        $inputDocument = [PdfSharp.Pdf.IO.PdfReader]::Open($file.FullName, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)
		# iterate over each page up to the total number of pages
		
        for ($index = 0; $index -lt $inputDocument.PageCount; $index++) {
		    # get a page
            $page = $inputDocument.Pages[$index]
            $output_page = $pdf.AddPage($page)
			
			# create the root bookmark on page 1, then the file name bookmark
			if($page_number -eq 1) {
			  # get the last part of directory name (https://social.technet.microsoft.com/Forums/Azure/en-US/21952986-bfd4-4858-a117-b9943f173c4b/list-last-folder-of-a-path?forum=ITCG)S
			  $srcDir = $file.DirectoryName.Split('\')[-1]
			  # gotcha: "true" has a $ in front
			  # https://devblogs.microsoft.com/powershell/boolean-values-and-operators/
			  $outline = $pdf.Outlines.Add($srcDir, $output_page, $true)
			  $index_page_name = [string]::Format("{0} - {1}", $file.Name, $page_number)
			  $outline.Outlines.Add($index_page_name, $output_page)
			}
			# otherwise, if this is a new file, just add to the existing outline
            if(($page_number -gt 1) -and ($index -eq 0)) {
			  $index_page_name = [string]::Format("{0} - {1}", $file.Name, $page_number)
			  $outline.Outlines.Add($index_page_name, $output_page)
			}

			# https://blogs.msdn.microsoft.com/walzenbach/2010/03/19/how-to-use-string-format-in-powershell/
			# write the short name of the file only
			# https://docs.microsoft.com/en-us/dotnet/api/system.io.fileinfo?view=netframework-4.7.2
			$footer_text = [string]::Format("{0} - {1}", $file.Name, $page_number)
			
			# Create a graphics object for this page. 
			# See http://www.pdfsharp.net/wiki/ConcatenateDocuments-sample.ashx
            $gfx = [PdfSharp.Drawing.XGraphics]::FromPdfPage($output_page, [PdfSharp.Drawing.XGraphicsPdfPageOptions]::Append)
			# Create the box in which the text will be drawn.
			$box = $page.MediaBox.ToXRect()
			$box.Inflate(0, -10)
			# Use the BottomCenter text format to put it on the bottom
			$gfx.DrawString($footer_text, $font, [PdfSharp.Drawing.XBrushes]::Black, $box, [PdfSharp.Drawing.XStringFormats]::BottomCenter)
			# increment the page number for adding the next page
			$page_number++
        }
    }

    Write-Host 'Saving PDF' $DestinationFile.FullName
	
	# delete the destination if set to force overwrite
    if((Test-Path $DestinationFile.FullName) -and $Force) { 
        Remove-Item $DestinationFile.FullName -Force 
    }
    $pdf.Save($DestinationFile.FullName)

    # if set to remove the source files, then delete the original
    if($RemoveSourceFiles) {
        Write-Host "Removing source files"
        foreach($file in $files) {
            Remove-Item -Path $file.FullName -Force:$Force
        }
    }
}
