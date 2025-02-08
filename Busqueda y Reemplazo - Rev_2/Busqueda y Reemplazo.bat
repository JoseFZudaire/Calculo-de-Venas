powershell Expand-Archive BusquedaYReemplazo-Rev_2.zip -DestinationPath %USERPROFILE%\Desktop\ -Force
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%userprofile%\Desktop\Busqueda y Reemplazo.lnk');$s.TargetPath='%userprofile%\Desktop\Busqueda y Reemplazo\AUTOCAD Redesign\bin\Debug\AUTOCAD Redesign.exe';$s.Save()"
echo "Ya esta listo para usarse el programa!"
pause