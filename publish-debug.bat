DEL AutoUpdaterTest-debug.zip
Copy .\App.debug.config .\bin\Debug\AutoUpdaterTest.exe.config /Y
"C:\Program Files\7-Zip\7z.exe" a -tzip AutoUpdaterTest-debug.zip^
 .\bin\Debug\*.exe^
 .\bin\Debug\*.dll^
 .\bin\Debug\*.pdb^
 
"C:\Program Files\7-Zip\7z.exe" d AutoUpdaterTest-debug.zip^
 AutoUpdaterTest.vshost.exe^
 