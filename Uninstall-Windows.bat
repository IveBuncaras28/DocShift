@echo off
powershell -NoProfile -Command "foreach($d in 'Desktop','Programs'){Remove-Item ([Environment]::GetFolderPath($d)+'\DocShift.lnk') -ErrorAction SilentlyContinue}"
rmdir /s /q "%LOCALAPPDATA%\DocShift"
echo DocShift removed.
pause
