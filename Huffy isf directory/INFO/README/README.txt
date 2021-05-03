README

SYNOPSIS: Compare Huffy daily loading report with ISF database and return new bookings in new bookings report.

INSTALL:
To install Compare-Huffy-isfs script, navigate to 'install' folder in Huffy isf directory -> INFO. Right click on
batch-converter.ps1, and select 'Run with Powershell'. A bat file of Compare-Huffy-isfs will be created which will
allow you double click and run from anywhere on your PC.

NOTES:
* Loading report must exist in Huffy isf directory during script runtime.
* Huffy isf directory must be saved onto the desktop.
* Huffy isf database is updated and saved at the end of Compare-Huffy-isf run time. A backup of the database
is also saved in Onedrive.
* If you cannot open up an excel file after running the script, open up the task manager and delete all
instances of Excel. Press ctrl + alt + del and navigate to task manager. Click on 'Details' tab and scroll down till 
you see instances titled Excel.exe. Right click on them and select 'End Task'. (For more help with the task manager, go 
to https://www.howtogeek.com/405806/windows-task-manager-the-complete-guide/).


AUTHOR: Dylan Rajca, 03.28.2021