# ps_pdfmerge
Merge PDFs under a directory using PowerShell, CoolOrange Module, and PDFSharp .NET extension

Installation / Configuration
============================
1. Download CoolOrange's PDFSharp DLL from:
https://support.coolorange.com/support/solutions/articles/22000211000-how-to-merge-multiple-pdf-with-pdfsharp-and-powershell
You will not need their cO.Pdfsharp.psm1 module (which provides the PdfMerge cmdlet)
unless you do not want the enhancements of a footer and bookmarking.

2. Modify this module and scripts to adjust the paths to where you installed the CoolOrange code.

3. Also modify the baseDir and destDir variables to control where you want to save the PDFs.

Note that there are also standalone test script for both the basic PdfMerge cmdlet
and the recursive version, which take no parameters. They would need all the variables modified.

Usage
=====

From Command Line
-----------------
1. After installation (see above), open cmd.exe

2. Run the following command:
`powershell.exe -ExecutionPolicy Bypass -Command "& 'C:\Users\USERNAME\Software\pdfMerge\pdfMergeRecurse.ps1'  -srcDir C:\pdfSrc -destDir C:\pdfDest"`
where the first part of the path is where you've installed.

To Use Module in Your Own Scripts
---------------------------------
1. After installation (see above), open PowerShell (command line or ISE).

2. Set to be allowed to run in PowerShell by first running the command below. This allows for scripts to be run in process scope.
See http://9to5it.com/bypass-the-powershell-execution-policy/

`Set-ExecutionPolicy Bypass -Scope Process`

You may need to run PowerShell as Administrator for this to work.

3. Load the module.


4. Then, set the variables as needed. Example:

~~~~
$srcDir = "C:\srcDir"
$destDir = "C:\destDir"
MergePDFRecurse $srcDir $destDir
~~~~

Note that if $destDir does not exist, it will be created.

Known Issues
============

* May not work with extremely long paths (need to test)

Todo
====

* Make the paths to the coolOrange code configurable as optional parameters. Otherwise, use defaults in source.
* Include a test script for the MergePDFRecurse which could be invoked from command line as a one liner. 
This would be useful for batch executions.
* Make it confirm whether the srcDir exists.
* Try/catch on errors
* Logging


