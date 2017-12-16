Param(
  [string]$target,
  [string]$shortcutbase
)

# Get the name of the last folder in $target
if ((Get-Item $target) -is [System.IO.DirectoryInfo])
{
# match 'test' in 'D:\Thomas\FastAcces 33\test'
$pattern = '(?:[\w]\:|\\)(?:\\[\w\-\ \.0-9]+)*(?:\\)([\w\-\ \.0-9]+)'
$target -match $pattern | Out-Null;
$shortcut_filename = $Matches[1];
}
else
{
# match 'test' in 'D:\Thomas\FastAcces 33\test\.lol'
$pattern = '(?:[\w]\:|\\)(?:\\[\w\-\ \.0-9]+)*(?:\\)([\w\-\ \.0-9]+)(?:\\)(?:[\w\-\.\ 0-9]+)*'
$target -match $pattern | Out-Null;
$shortcut_filename = $Matches[1];
# match 'D:\Thomas\FastAcces 33\test\' in 'D:\Thomas\FastAcces 33\test\.lol'
$pattern = '((?:[\w]\:|\\)(?:\\[\w\-\ \.0-9]+)*)(?:\\)(?:[\w\-\ \.0-9]+)'
$target -match $pattern | Out-Null;
$target = $Matches[1];
}

# Add Shortcut in the $target folder with the name from above
$ws = New-Object -ComObject WScript.Shell; 
$s = $ws.CreateShortcut($shortcutbase + '\' + $shortcut_filename + '.lnk'); 
$S.TargetPath = $target; 
$S.Save()