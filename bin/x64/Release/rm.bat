del /Q /S *.bak
del /Q /S *.iobj
del /Q /S *.ipdb
del /Q /S *.pdb
rem del /Q /S *extract.docx
rem del /Q /S *out.docx

rmdir /Q /S __pycache__
rmdir /Q /S build
rmdir /Q /S dist
rmdir /Q /S tmp\split
rmdir /Q /S tmp\split-processed

pause