# Petty_scripts_for_tests
Petty but useful scripts I used to run some tests over codes

## compare_xlsx
### Interest
This script reads two csv files and a "maximum range letter" (as identification of an Excel column). It converts the csv files to xlsx files and compares them line by line, cell by cell until the indicated max range. It then modifies the xlsx files by changing the background color of the cells: green if they match, red if not.

### Usage (with Windows)
Run the compare-csv-data-and-export.cmd command file to run the script. If needed, change the setting (location of the csv files for example) inside the .cmd file before running it
