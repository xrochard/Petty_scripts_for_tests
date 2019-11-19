@echo off

set WORKING_DIR=C:\Github\Petty_scripts_for_tests\\temp
set RESOURCES_DIR=C:\Github\Petty_scripts_for_tests\src

echo "WORKING_DIR is: " %WORKING_DIR%
set /p max_range="What is the maximum range? "

python %RESOURCES_DIR%/Compare_xlsx.py --actual %WORKING_DIR%/data.csv --expected %WORKING_DIR%/export.csv --max_range %max_range%

echo.
echo -----------------------------------
echo End of script
echo.

pause


