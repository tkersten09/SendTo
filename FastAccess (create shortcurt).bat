set TARGET="%1"
set SHORTCUTBASE="D:\Thomas\FastAccess"
set PWS=powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile

%PWS% -f "C:\Users\Thomas\AppData\Roaming\Microsoft\Windows\SendTo\FastAccess.ps1" -target %TARGET% -shortcutbase %SHORTCUTBASE%
