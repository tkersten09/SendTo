Param(
  [string]$target,
  [string]$shortcutbase
)

if ((Get-Item $target) -is [System.IO.DirectoryInfo])
{
# (?:[\w]\:|\\)(?:\\[\w\-\ \.0-9]+)*(?:\\)([\w\-\ \.0-9]+)
$pattern = '(?:[\w]\:|\\)(?:\\[\w\-\ \.0-9]+)*(?:\\)([\w\-\ \.0-9]+)'
$target -match $pattern | Out-Null;
$shortcut = $Matches[1];
}
else
{
# match 'test' in 'D:\Thomas\FastAcces 33\test\.lol'
# (?:[\w]\:|\\)(?:\\[\w\-\ \.0-9]+)*(?:\\)([\w\-\ \.0-9]+)(?:\\)(?:[\w\-\.\ 0-9]+)*
$pattern = '(?:[\w]\:|\\)(?:\\[\w\-\ \.0-9]+)*(?:\\)([\w\-\ \.0-9]+)(?:\\)(?:[\w\-\.\ 0-9]+)*'
$target -match $pattern | Out-Null;
$shortcut = $Matches[1];
# match 'D:\Thomas\FastAcces 33\test\' in 'D:\Thomas\FastAcces 33\test\.lol'
$pattern = '((?:[\w]\:|\\)(?:\\[\w\-\ \.0-9]+)*)(?:\\)(?:[\w\-\ \.0-9]+)'
$target -match $pattern | Out-Null;
$target = $Matches[1];
}

$ws = New-Object -ComObject WScript.Shell; 
$s = $ws.CreateShortcut($shortcutbase + '\' + $shortcut + '.lnk'); 
$S.TargetPath = $target; 
$S.Save()