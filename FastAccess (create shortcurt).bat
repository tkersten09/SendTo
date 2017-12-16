set TARGET="%1"
set SHORTCUTBASE="D:\Thomas\FastAccess"
set PWS=powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile

:: Run Powershell in cmd https://stackoverflow.com/questions/16436405/how-to-run-powershell-in-cmd
%PWS% -f "C:\Users\Thomas\AppData\Roaming\Microsoft\Windows\SendTo\FastAccess.ps1" -target %TARGET% -shortcutbase %SHORTCUTBASE%
