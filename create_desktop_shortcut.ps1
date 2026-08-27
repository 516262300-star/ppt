$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$appScript = Join-Path $scriptDir "app.ps1"
$desktopDir = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktopDir "设计师型号月报助手.lnk"

if (-not (Test-Path -LiteralPath $appScript)) {
    throw "找不到启动程序：$appScript"
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$shortcut.Arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$appScript`""
$shortcut.WorkingDirectory = $scriptDir

$customIcon = Join-Path $scriptDir "assets\月报助手.ico"
$powerPoint = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
if (Test-Path -LiteralPath $customIcon) {
    $shortcut.IconLocation = "$customIcon,0"
} elseif (Test-Path -LiteralPath $powerPoint) {
    $shortcut.IconLocation = "$powerPoint,0"
} else {
    $shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,261"
}

$shortcut.Description = "录入作品和交易数据，生成一份完整美工月会 PPT"
$shortcut.Save()
Write-Output $shortcutPath
