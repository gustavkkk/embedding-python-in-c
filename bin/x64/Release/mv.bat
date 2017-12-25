del /Q /S *.bak
del /Q /S *.iobj
del /Q /S *.ipdb
del /Q /S *.pdb

rmdir /Q /S __pycache__

rem set target="split"
xcopy db\* .\dist\%target%\db
xcopy tmp\* .\dist\%target%\tmp /s /i
echo F| xcopy dict.txt .\dist\%target%\dict.txt
echo F| xcopy idf.txt .\dist\%target%\idf.txt

rem xcopy db\* .\dist\extract\db
rem xcopy tmp\* .\dist\extract\tmp /s /i
rem echo F| xcopy dict.txt .\dist\extract\dict.txt
rem echo F| xcopy idf.txt .\dist\extract\idf.txt

pause