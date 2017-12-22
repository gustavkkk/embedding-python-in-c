del /Q /S *.aps
del /Q /S *.sdf
del /Q /S *.suo
del /Q /S *.pdb
del /Q /S *.ilk
del /Q /S *.exp

rmdir /Q /S ipch
rmdir /Q /S TestBinding\Win32
rmdir /Q /S TestBinding\x64
rmdir /Q /S Release
rmdir /Q /S Debug
rmdir /Q /S .vs
rmdir /Q /S TestBinding\GeneratedFiles

pause