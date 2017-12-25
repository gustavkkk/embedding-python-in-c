del /Q /S *.aps
del /Q /S *.sdf
del /Q /S *.suo
del /Q /S *.pdb
del /Q /S *.ilk
del /Q /S *.exp
del /Q /S *.bak

rmdir /Q /S ipch
rmdir /Q /S TestBinding\Win32
rmdir /Q /S TestBinding\x64
rmdir /Q /S MoreComplex\Win32
rmdir /Q /S MoreComplex\x64
rmdir /Q /S IPC\Win32
rmdir /Q /S IPC\x64
rmdir /Q /S Release
rmdir /Q /S Debug
rmdir /Q /S .vs
rmdir /Q /S TestBinding\GeneratedFiles
rmdir /Q /S MoreComplex\GeneratedFiles
rmdir /Q /S IPC\GeneratedFiles
rmdir /Q /S TestBinding\tmp
rmdir /Q /S MoreComplex\tmp
rmdir /Q /S IPC\tmp

pause