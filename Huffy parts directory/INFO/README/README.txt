README

SYNOPSIS: Grab parts data from input.xlsx and parse into a table that lists tariffs with totals (output.txt)

INSTALL:
To install script, navigate to 'install' folder in INFO directory. Right click on
batch-converter.ps1, and select 'Run with Powershell'. A bat file of the script will be created which will
allow you double click and run from anywhere on your PC.

NOTES:
* Huffy parts directory folder must exist on the desktop.

* input.xlsx must contain "Hts Numbers", "Qty" and "Cost" in the header. These cells are locked to prevent changes.
	- Password to unprotect workbook: "password"

* The script will not execute properly if there are spaces in any of the input files.
	- If there are spaces on the invoice, write in "blank" as a placeholder in the input file. 
	- If there is a space where an hts code should be, add in the correct hts code or write in "blank"
	as a placeholder in the input file.

* Always double check to make sure output is accurate (check totals, add up values etc). Adobe does not
always copy data accurately.

AUTHOR: Dylan Rajca, 04.30.2021