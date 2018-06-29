# ps_pdfmerge
Merge PDFs using PowerShell, CoolOrange Module, and PDFSharp .NET extension

Installation / Configuration
----------------------------
1. Download CoolOrange's MergePDF PowerShell module and accompanying PDFSharp DLL from:
https://support.coolorange.com/support/solutions/articles/22000211000-how-to-merge-multiple-pdf-with-pdfsharp-and-powershell

2. Modify the script for where you installed the CoolOrange code.

3. Also modify the baseDir and destDir variables to control where you want to save the PDFs.

In a future version, I may turn this into a module so you could simply pass those in as positional parameters, so no edits are needed.

Usage
-----
To run this, you need to set it to be allowed to run in PS ISE by first running the command below in a PowerShell tab -
allows for scripts to run, in process scope Set-ExecutionPolicy Bypass -Scope Process.

Then you can run it from PowerShell ISE.


