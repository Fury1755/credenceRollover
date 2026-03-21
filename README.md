# credenceRollover
code for rolling over financial statements in credence advisory

INSTALLATION

First you need to activate your "Personal Macro Workbook", which lets you save and use the macro for any Excel workbook. Otherwise you have to keep
installing the code for every Excel file you open and want to use it on.

To activate Personal Macro Workbook:
Open excel -> right click on top ribbon -> select customize ribbon -> click developer
select developer on top -> select record macro -> select Store Macro in "Personal Macro Workbook"

now download the .bas and .cls files from here. Click on "Code" and download the autoRollover as zip.
Unzip and extract to a folder; should have a few files with .bas and .cls in there.
Open Excel - > Developer -> Macro -> Visual Basic
Look at top left area, You should see a project ending in "(PERSONAL.XLSB)" (that's your personal macro workbook). If you don't, press Ctrl + R to 
show the project explorer and you should see it.
Go to File -> import file
select and import the extracted files (one by one? I don't know, but there aren't that many)
You should see them under your PERSONAL.XLSB
Installation complete!

USAGE
create a new copy of previous year working
In Excel -> Developer -> Macros -> InitializeRolloverWorkbooks
File explorer popup -> select the previous year working
File explorer popup again -> select the new copy
If it doesn't work the first time (it probably won't), run it again

WARNINGS
My memory usage increases by 3-7 MB everytime I run the macro. Don't worry, it will reset after you close all Excel tabs.

HOW TO UPDATE
The macro is currently programmed to run for FY2024 to FY2025. If you are rolling over for FY2025-FY2026, just 
Excel -> Developer -> Macros -> select InitializeRolloverWorkbooks -> Edit Macro -> Ctrl + F -> Replace -> select "Current Project"
-> then just do your find and replace from highest year to lowest year. Don't replace 2024 with 2025, then replace 2025 with 2026 (your original "2024" is now "2026")

