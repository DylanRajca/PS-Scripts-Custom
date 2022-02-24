README

SYNOPSIS: Grab parts data from input.xlsx and parse into a table that lists tariffs with totals (output.txt)

INSTALL:
To install script, navigate to 'install' folder in INFO directory. Right click on
batch-converter.ps1, and select 'Run with Powershell'. A bat file of the script will be created which will
allow you double click and run from anywhere on your PC.

NOTES:
* Working directory folder must exist on the desktop or in documents.

* input.xlsx must contain "CI-HTS-CODES", "Qty", "Cost", "PL-HTS-CODES", "NET", "Gross" in the header. These cells are locked to prevent changes.
	- Password to unprotect workbook: "password"

* Always double check to make sure output is accurate (check totals, add up values etc). 

AUTHOR: Dylan Rajca, 01/26/2022