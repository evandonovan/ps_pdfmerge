# ps_pdfmerge
Merge PDFs under a directory using PowerShell, CoolOrange Module, and PDFSharp .NET extension

Installation / Configuration
----------------------------
1. Download CoolOrange's MergePDF PowerShell module and accompanying PDFSharp DLL from:
https://support.coolorange.com/support/solutions/articles/22000211000-how-to-merge-multiple-pdf-with-pdfsharp-and-powershell

2. Modify the module to adjust the paths to where you installed the CoolOrange code.

3. Also modify the baseDir and destDir variables to control where you want to save the PDFs.

Note that there is also a standalone test script, which takes no parameters. It would need all the variables modified.

Usage
-----

1. After installation (see above), open PowerShell (command line or ISE).

2. Set to be allowed to run in PowerShell by first running the command below. This allows for scripts to be run in process scope.
See http://9to5it.com/bypass-the-powershell-execution-policy/

Set-ExecutionPolicy Bypass -Scope Process

3. Then, set the variables as needed. Example:

~~~~
$srcDir = "C:\srcDir"
$destDir = "C:\destDir"
MergePDFRecurse $srcDir $destDir
~~~~

Note that if $destDir does not exist, it will be created.

Known Issues
------------

* May not work with extremely long paths (need to test)
* Some underlying code (not mine) calls an AddLog cmdlet that may not exist on your machine. It will throw an error per PDF to the PS command line. This is harmless, I think - probably just happens if you don't have Visual Studio.

Todo
----

* Make the paths to the coolOrange code configurable as optional parameters. Otherwise, use defaults in source.
* Include a test script for the MergePDFRecurse which could be invoked from command line as a one liner. 
This would be useful for batch executions.
* Make it confirm whether the srcDir exists.
* Try/catch on errors
* Logging


