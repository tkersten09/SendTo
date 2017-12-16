set TARGET="%1"
set SHORTCUTBASE="D:\Thomas\FastAccess"
set PWS=powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile

:: %PWS% -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut(%SHORTCUT%); $S.TargetPath = %TARGET%; $S.Save()"
%PWS% -f "C:\Users\Thomas\AppData\Roaming\Microsoft\Windows\SendTo\FastAccess.ps1" -target %TARGET% -shortcutbase %SHORTCUTBASE%

:: wait for development
:: timeout 100