param(
    [switch]$SmokeTest,
    [switch]$SmokeGenerate,
    [string]$TestWorksRoot = "",
    [switch]$SmokeClipboard,
    [string]$TestImagePath = ""
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$runScript = Join-Path $scriptDir "run_portfolio.ps1"
$settingsDir = Join-Path $scriptDir ".build"
$settingsFile = Join-Path $settingsDir "portfolio-settings.json"
$sourceFile = Join-Path $settingsDir "portfolio-source.txt"
$desktopDir = [Environment]::GetFolderPath("Desktop")
$clipboardRoot = Join-Path $settingsDir $(if ($SmokeClipboard) { "clipboard-smoke" } else { "clipboard-portfolio" })
New-Item -ItemType Directory -Force -Path $settingsDir,$clipboardRoot | Out-Null
$script:worksRoot = $clipboardRoot
$script:usingClipboard = $true
if (Test-Path -LiteralPath $sourceFile -PathType Leaf) {
    $savedSource = (Get-Content -Raw -LiteralPath $sourceFile).Trim()
    if ($savedSource -and (Test-Path -LiteralPath $savedSource -PathType Container)) {
        $script:worksRoot = $savedSource
        $script:usingClipboard = ([System.IO.Path]::GetFullPath($savedSource) -eq [System.IO.Path]::GetFullPath($clipboardRoot))
    }
}

function Safe-Period([string]$period) { return ($period -replace '[\\/:*?<>|]', '-') }
function Default-Output([string]$period) { return (Join-Path $desktopDir "美工与设计师作品展示_$(Safe-Period $period).pptx") }
function Quote-ProcessArgument([string]$value) { return '"' + $value.Replace('"', '\"') + '"' }
function Save-Settings($values) { $values | ConvertTo-Json | Set-Content -LiteralPath $settingsFile -Encoding UTF8 }
function Safe-Name([string]$value) {
    $result = $value.Trim()
    foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) { $result = $result.Replace([string]$char, '_') }
    return $result.Trim().TrimEnd('.')
}
function Is-ImageFile([string]$filePath) {
    return @('.png','.jpg','.jpeg','.webp') -contains [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
}
function Unique-ImagePath([string]$folder, [string]$extension) {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $token = [guid]::NewGuid().ToString('N').Substring(0, 6)
    return (Join-Path $folder "${stamp}_${token}${extension}")
}

$defaultPeriod = Get-Date -Format "yyyy年M月"
$defaults = [ordered]@{ Period = $defaultPeriod; OutputFile = Default-Output $defaultPeriod; OpenAfter = $true }
if (Test-Path -LiteralPath $settingsFile) {
    try {
        $saved = Get-Content -Raw -LiteralPath $settingsFile | ConvertFrom-Json
        foreach ($name in @('Period', 'OutputFile', 'OpenAfter')) {
            if ($null -ne $saved.$name -and "$($saved.$name)" -ne "") { $defaults[$name] = $saved.$name }
        }
    } catch {
        # 设置损坏时回退到默认值。
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "美工与设计师作品展示 PPT"
$form.StartPosition = "CenterScreen"
$form.ClientSize = New-Object System.Drawing.Size(900, 730)
$form.MinimumSize = New-Object System.Drawing.Size(916, 769)
$form.BackColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
$form.KeyPreview = $true
$appIcon = Join-Path $scriptDir "assets\月报助手.ico"
if (Test-Path -LiteralPath $appIcon) { $form.Icon = New-Object System.Drawing.Icon -ArgumentList $appIcon }

$header = New-Object System.Windows.Forms.Panel
$header.Location = New-Object System.Drawing.Point(0, 0)
$header.Size = New-Object System.Drawing.Size(900, 84)
$header.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F26B38")
$header.Anchor = "Top,Left,Right"
$form.Controls.Add($header)

$title = New-Object System.Windows.Forms.Label
$title.Text = "美工与设计师作品展示 PPT"
$title.ForeColor = [System.Drawing.Color]::White
$title.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 18, [System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(26, 14)
$header.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "复制图片后直接按 Ctrl+V，软件自动按身份和姓名归集"
$subtitle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#FFF0E9")
$subtitle.AutoSize = $true
$subtitle.Location = New-Object System.Drawing.Point(29, 52)
$header.Controls.Add($subtitle)

$roleLabel = New-Object System.Windows.Forms.Label
$roleLabel.Text = "1. 身份"
$roleLabel.Location = New-Object System.Drawing.Point(30, 116)
$roleLabel.Size = New-Object System.Drawing.Size(80, 28)
$form.Controls.Add($roleLabel)
$roleBox = New-Object System.Windows.Forms.ComboBox
$roleBox.DropDownStyle = "DropDownList"
$null = $roleBox.Items.Add("美工")
$null = $roleBox.Items.Add("设计师")
$roleBox.SelectedIndex = 0
$roleBox.Location = New-Object System.Drawing.Point(110, 110)
$roleBox.Size = New-Object System.Drawing.Size(130, 30)
$form.Controls.Add($roleBox)

$nameLabel = New-Object System.Windows.Forms.Label
$nameLabel.Text = "2. 姓名"
$nameLabel.Location = New-Object System.Drawing.Point(270, 116)
$nameLabel.Size = New-Object System.Drawing.Size(80, 28)
$form.Controls.Add($nameLabel)
$nameBox = New-Object System.Windows.Forms.TextBox
$nameBox.Location = New-Object System.Drawing.Point(350, 110)
$nameBox.Size = New-Object System.Drawing.Size(190, 30)
$form.Controls.Add($nameBox)

$pasteButton = New-Object System.Windows.Forms.Button
$pasteButton.Text = "粘贴图片  Ctrl+V"
$pasteButton.Location = New-Object System.Drawing.Point(570, 103)
$pasteButton.Size = New-Object System.Drawing.Size(290, 42)
$pasteButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F26B38")
$pasteButton.ForeColor = [System.Drawing.Color]::White
$pasteButton.FlatStyle = "Flat"
$pasteButton.FlatAppearance.BorderSize = 0
$pasteButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($pasteButton)

$pasteHelp = New-Object System.Windows.Forms.Label
$pasteHelp.Text = "先选择身份、填写姓名，再从微信、网页或其他软件复制图片；可重复粘贴，也支持一次复制多张图片文件。"
$pasteHelp.Location = New-Object System.Drawing.Point(30, 155)
$pasteHelp.Size = New-Object System.Drawing.Size(830, 28)
$pasteHelp.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#657086")
$form.Controls.Add($pasteHelp)

$chooseImagesButton = New-Object System.Windows.Forms.Button
$chooseImagesButton.Text = "选择图片文件..."
$chooseImagesButton.Location = New-Object System.Drawing.Point(110, 190)
$chooseImagesButton.Size = New-Object System.Drawing.Size(170, 34)
$form.Controls.Add($chooseImagesButton)
$importFolderButton = New-Object System.Windows.Forms.Button
$importFolderButton.Text = "从现有文件夹导入..."
$importFolderButton.Location = New-Object System.Drawing.Point(300, 190)
$importFolderButton.Size = New-Object System.Drawing.Size(190, 34)
$form.Controls.Add($importFolderButton)
$openFolderButton = New-Object System.Windows.Forms.Button
$openFolderButton.Text = "打开当前图片目录"
$openFolderButton.Location = New-Object System.Drawing.Point(510, 190)
$openFolderButton.Size = New-Object System.Drawing.Size(180, 34)
$form.Controls.Add($openFolderButton)

$sourceLabel = New-Object System.Windows.Forms.Label
$sourceLabel.Text = if ($script:usingClipboard) { "当前来源：直接粘贴图片" } else { "当前来源：已有文件夹" }
$sourceLabel.Location = New-Object System.Drawing.Point(710, 197)
$sourceLabel.Size = New-Object System.Drawing.Size(160, 24)
$sourceLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#24489E")
$form.Controls.Add($sourceLabel)

$collectionList = New-Object System.Windows.Forms.ListView
$collectionList.Location = New-Object System.Drawing.Point(30, 238)
$collectionList.Size = New-Object System.Drawing.Size(840, 165)
$collectionList.View = "Details"
$collectionList.FullRowSelect = $true
$collectionList.GridLines = $true
$null = $collectionList.Columns.Add("身份", 120)
$null = $collectionList.Columns.Add("姓名", 260)
$null = $collectionList.Columns.Add("图片数量", 120)
$null = $collectionList.Columns.Add("图片位置", 330)
$form.Controls.Add($collectionList)

$deletePersonButton = New-Object System.Windows.Forms.Button
$deletePersonButton.Text = "删除选中人员图片"
$deletePersonButton.Location = New-Object System.Drawing.Point(30, 413)
$deletePersonButton.Size = New-Object System.Drawing.Size(180, 34)
$form.Controls.Add($deletePersonButton)
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "清空本次粘贴"
$clearButton.Location = New-Object System.Drawing.Point(225, 413)
$clearButton.Size = New-Object System.Drawing.Size(160, 34)
$form.Controls.Add($clearButton)

$periodLabel = New-Object System.Windows.Forms.Label
$periodLabel.Text = "3. 报告月份"
$periodLabel.Location = New-Object System.Drawing.Point(430, 420)
$periodLabel.Size = New-Object System.Drawing.Size(120, 28)
$form.Controls.Add($periodLabel)
$periodBox = New-Object System.Windows.Forms.TextBox
$periodBox.Text = $defaults.Period
$periodBox.Location = New-Object System.Drawing.Point(550, 413)
$periodBox.Size = New-Object System.Drawing.Size(180, 30)
$form.Controls.Add($periodBox)

$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Text = "4. 输出 PPT"
$outputLabel.Location = New-Object System.Drawing.Point(30, 475)
$outputLabel.Size = New-Object System.Drawing.Size(120, 28)
$form.Controls.Add($outputLabel)
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Text = $defaults.OutputFile
$outputBox.Location = New-Object System.Drawing.Point(150, 468)
$outputBox.Size = New-Object System.Drawing.Size(600, 30)
$outputBox.Anchor = "Top,Left,Right"
$form.Controls.Add($outputBox)
$outputButton = New-Object System.Windows.Forms.Button
$outputButton.Text = "保存为..."
$outputButton.Location = New-Object System.Drawing.Point(765, 467)
$outputButton.Size = New-Object System.Drawing.Size(105, 32)
$outputButton.Anchor = "Top,Right"
$form.Controls.Add($outputButton)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.ReadOnly = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.Location = New-Object System.Drawing.Point(30, 520)
$statusBox.Size = New-Object System.Drawing.Size(840, 100)
$statusBox.Anchor = "Top,Bottom,Left,Right"
$statusBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#FFF9F6")
$statusBox.Text = "等待粘贴图片。"
$form.Controls.Add($statusBox)

$openAfterCheck = New-Object System.Windows.Forms.CheckBox
$openAfterCheck.Text = "完成后自动打开 PPT"
$openAfterCheck.Checked = [bool]$defaults.OpenAfter
$openAfterCheck.Location = New-Object System.Drawing.Point(30, 660)
$openAfterCheck.Size = New-Object System.Drawing.Size(190, 28)
$openAfterCheck.Anchor = "Bottom,Left"
$form.Controls.Add($openAfterCheck)

$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = "完成录入并返回"
$generateButton.Location = New-Object System.Drawing.Point(630, 645)
$generateButton.Size = New-Object System.Drawing.Size(240, 48)
$generateButton.Anchor = "Bottom,Right"
$generateButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F26B38")
$generateButton.ForeColor = [System.Drawing.Color]::White
$generateButton.FlatStyle = "Flat"
$generateButton.FlatAppearance.BorderSize = 0
$generateButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($generateButton)

if (-not $SmokeGenerate) {
    foreach ($control in @($periodLabel, $periodBox, $outputLabel, $outputBox, $outputButton, $openAfterCheck)) {
        $control.Visible = $false
    }
    $statusBox.Location = New-Object System.Drawing.Point(30, 468)
    $statusBox.Size = New-Object System.Drawing.Size(840, 152)
    $statusBox.Text = "作品会保存到本次月会内容中。录入完成后返回主窗口，再生成一份完整月会 PPT。"
}

$imageDialog = New-Object System.Windows.Forms.OpenFileDialog
$imageDialog.Filter = "图片文件 (*.png;*.jpg;*.jpeg;*.webp)|*.png;*.jpg;*.jpeg;*.webp|所有文件 (*.*)|*.*"
$imageDialog.Multiselect = $true
$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$saveDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveDialog.Filter = "PowerPoint 文件 (*.pptx)|*.pptx"
$saveDialog.AddExtension = $true

function Current-PersonDirectory {
    $role = Safe-Name "$($roleBox.SelectedItem)"
    $name = Safe-Name $nameBox.Text
    if (-not $role -or -not $name) { return $null }
    return (Join-Path $clipboardRoot (Join-Path $role $name))
}

function Refresh-Collection {
    $collectionList.Items.Clear()
    $root = $script:worksRoot
    if (-not (Test-Path -LiteralPath $root)) { return }
    $groups = @{}
    $rootFull = [System.IO.Path]::GetFullPath($root).TrimEnd('\')
    foreach ($file in Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue) {
        if (-not (Is-ImageFile $file.FullName)) { continue }
        $relative = $file.FullName.Substring($rootFull.Length).TrimStart('\')
        $parts = $relative -split '\\'
        if ($parts.Count -lt 2) { continue }
        if ($parts.Count -ge 3 -and @('美工','设计师') -contains $parts[0]) {
            $role = $parts[0]; $name = $parts[1]; $folder = Join-Path $root (Join-Path $role $name)
        } else {
            $role = ''; $name = $parts[0]; $folder = Join-Path $root $name
        }
        $key = "$role|$name|$folder"
        if (-not $groups.ContainsKey($key)) { $groups[$key] = 0 }
        $groups[$key]++
    }
    foreach ($key in ($groups.Keys | Sort-Object)) {
        $parts = $key -split '\|', 3
        $item = New-Object System.Windows.Forms.ListViewItem($(if ($parts[0]) { $parts[0] } else { '未分类' }))
        $null = $item.SubItems.Add($parts[1])
        $null = $item.SubItems.Add("$($groups[$key])")
        $null = $item.SubItems.Add($parts[2])
        $item.Tag = $parts[2]
        $null = $collectionList.Items.Add($item)
    }
    $deletePersonButton.Enabled = $script:usingClipboard
    $clearButton.Enabled = $script:usingClipboard
}

function Import-ImageFiles([string[]]$files) {
    $personDir = Current-PersonDirectory
    if (-not $personDir) {
        if (-not ($SmokeTest -or $SmokeClipboard)) { [System.Windows.Forms.MessageBox]::Show("请先选择身份并填写姓名。", "提示", "OK", "Warning") }
        return 0
    }
    New-Item -ItemType Directory -Force -Path $personDir | Out-Null
    $count = 0
    foreach ($file in $files) {
        if (-not (Test-Path -LiteralPath $file -PathType Leaf) -or -not (Is-ImageFile $file)) { continue }
        $destination = Unique-ImagePath $personDir ([System.IO.Path]::GetExtension($file).ToLowerInvariant())
        Copy-Item -LiteralPath $file -Destination $destination
        $count++
    }
    $script:worksRoot = $clipboardRoot
    $script:usingClipboard = $true
    $clipboardRoot | Set-Content -LiteralPath $sourceFile -Encoding UTF8
    $sourceLabel.Text = "当前来源：直接粘贴图片"
    Refresh-Collection
    return $count
}

function Paste-ClipboardImages($dataObject = $null) {
    if ($null -eq $dataObject) { $dataObject = [System.Windows.Forms.Clipboard]::GetDataObject() }
    $personDir = Current-PersonDirectory
    if (-not $personDir) {
        if (-not $SmokeClipboard) { [System.Windows.Forms.MessageBox]::Show("请先选择身份并填写姓名。", "提示", "OK", "Warning") }
        return 0
    }
    New-Item -ItemType Directory -Force -Path $personDir | Out-Null
    $count = 0
    if ($dataObject -and $dataObject.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $count = Import-ImageFiles ([string[]]$dataObject.GetData([System.Windows.Forms.DataFormats]::FileDrop))
    } elseif ($dataObject -and $dataObject.GetDataPresent([System.Windows.Forms.DataFormats]::Bitmap)) {
        $image = $dataObject.GetData([System.Windows.Forms.DataFormats]::Bitmap)
        if ($image) {
            $destination = Unique-ImagePath $personDir '.png'
            $image.Save($destination, [System.Drawing.Imaging.ImageFormat]::Png)
            $count = 1
            $script:worksRoot = $clipboardRoot
            $script:usingClipboard = $true
            $clipboardRoot | Set-Content -LiteralPath $sourceFile -Encoding UTF8
            $sourceLabel.Text = "当前来源：直接粘贴图片"
            Refresh-Collection
        }
    }
    if ($count -eq 0 -and -not $SmokeClipboard) {
        [System.Windows.Forms.MessageBox]::Show("剪贴板中没有可用图片。请先复制图片或图片文件。", "提示", "OK", "Information")
    } elseif ($count -gt 0) {
        $statusBox.Text = "已为 $($roleBox.SelectedItem) $($nameBox.Text.Trim()) 添加 $count 张图片。可继续复制并粘贴。"
    }
    return $count
}

$pasteButton.Add_Click({ [void](Paste-ClipboardImages) })
$form.Add_KeyDown({
    param($sender, $eventArgs)
    if ($eventArgs.Control -and $eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::V) {
        [void](Paste-ClipboardImages)
        $eventArgs.SuppressKeyPress = $true
    }
})
$chooseImagesButton.Add_Click({
    if ($imageDialog.ShowDialog() -eq "OK") {
        $count = Import-ImageFiles $imageDialog.FileNames
        if ($count -gt 0) { $statusBox.Text = "已导入 $count 张图片。" }
    }
})
$importFolderButton.Add_Click({
    $folderDialog.Description = "选择包含姓名文件夹的作品图片根目录"
    if ($folderDialog.ShowDialog() -eq "OK") {
        $script:worksRoot = $folderDialog.SelectedPath
        $script:usingClipboard = $false
        $script:worksRoot | Set-Content -LiteralPath $sourceFile -Encoding UTF8
        $sourceLabel.Text = "当前来源：已有文件夹"
        Refresh-Collection
        $statusBox.Text = "已读取作品文件夹：$($script:worksRoot)"
    }
})
$openFolderButton.Add_Click({
    if (-not (Test-Path -LiteralPath $script:worksRoot)) { New-Item -ItemType Directory -Force -Path $script:worksRoot | Out-Null }
    Start-Process -FilePath "explorer.exe" -ArgumentList (Quote-ProcessArgument $script:worksRoot)
})
$deletePersonButton.Add_Click({
    if (-not $script:usingClipboard -or $collectionList.SelectedItems.Count -eq 0) { return }
    $folder = "$($collectionList.SelectedItems[0].Tag)"
    $rootFull = [System.IO.Path]::GetFullPath($clipboardRoot).TrimEnd('\') + '\'
    $folderFull = [System.IO.Path]::GetFullPath($folder)
    if (-not $folderFull.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) { return }
    if ([System.Windows.Forms.MessageBox]::Show("确认删除选中人员本次粘贴的全部图片？", "确认", "YesNo", "Question") -eq "Yes") {
        if (Test-Path -LiteralPath $folderFull) { Remove-Item -LiteralPath $folderFull -Recurse -Force }
        Refresh-Collection
    }
})
$clearButton.Add_Click({
    if (-not $script:usingClipboard) { return }
    if ([System.Windows.Forms.MessageBox]::Show("确认清空本次粘贴的全部作品图片？", "确认", "YesNo", "Question") -eq "Yes") {
        $resolved = [System.IO.Path]::GetFullPath($clipboardRoot)
        $settingsResolved = [System.IO.Path]::GetFullPath($settingsDir).TrimEnd('\') + '\'
        if (-not $resolved.StartsWith($settingsResolved, [StringComparison]::OrdinalIgnoreCase)) { throw "拒绝清理意外目录：$resolved" }
        if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Recurse -Force }
        New-Item -ItemType Directory -Force -Path $resolved | Out-Null
        Refresh-Collection
        $statusBox.Text = "本次粘贴图片已清空。"
    }
})
$periodBox.Add_Leave({ if ($periodBox.Text.Trim()) { $outputBox.Text = Default-Output $periodBox.Text.Trim() } })
$outputButton.Add_Click({
    $saveDialog.FileName = [System.IO.Path]::GetFileName($outputBox.Text)
    $parent = Split-Path -Parent $outputBox.Text
    if ($parent -and (Test-Path -LiteralPath $parent)) { $saveDialog.InitialDirectory = $parent }
    if ($saveDialog.ShowDialog() -eq "OK") { $outputBox.Text = $saveDialog.FileName }
})

$script:activeProcess = $null
$script:activeOptions = $null
$script:smokeResult = $null
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500
$timer.Add_Tick({
    if (-not $script:activeProcess) { $timer.Stop(); return }
    if (-not $script:activeProcess.HasExited) { return }
    $timer.Stop()
    $process = $script:activeProcess
    $options = $script:activeOptions
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $exitCode = $process.ExitCode
    $process.Dispose()
    $script:activeProcess = $null
    $script:activeOptions = $null
    $generateButton.Enabled = $true
    $generateButton.Text = "生成作品展示 PPT"
    $statusBox.Text = ($stdout + "`r`n" + $stderr).Trim()
    $success = ($exitCode -eq 0 -and (Test-Path -LiteralPath $options.OutputFile))
    if ($SmokeGenerate) {
        $script:smokeResult = [pscustomobject]@{ Success = $success; ExitCode = $exitCode; OutputFile = $options.OutputFile; Log = $statusBox.Text }
        $form.Close()
        return
    }
    if ($success) {
        [System.Windows.Forms.MessageBox]::Show("作品展示 PPT 已生成：`r`n$($options.OutputFile)", "完成", "OK", "Information")
        if ($options.OpenAfter) { Start-Process -FilePath $options.OutputFile }
    } else {
        [System.Windows.Forms.MessageBox]::Show("生成失败，请查看窗口中的运行日志。", "生成失败", "OK", "Error")
    }
})

$generateButton.Add_Click({
    if (-not $SmokeGenerate) {
        if ($collectionList.Items.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("请先粘贴图片或导入作品文件夹。", "提示", "OK", "Warning")
            return
        }
        $script:worksRoot | Set-Content -LiteralPath $sourceFile -Encoding UTF8
        $form.Close()
        return
    }
    $period = $periodBox.Text.Trim()
    $output = $outputBox.Text.Trim()
    $problems = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $script:worksRoot -PathType Container)) { $problems.Add("请先粘贴图片或导入作品文件夹。") }
    if ($collectionList.Items.Count -eq 0) { $problems.Add("当前没有可生成的作品图片。") }
    if (-not $period) { $problems.Add("请填写报告月份。") }
    if (-not $output -or [System.IO.Path]::GetExtension($output).ToLowerInvariant() -ne ".pptx") { $problems.Add("请选择 .pptx 输出文件。") }
    if ($problems.Count -gt 0) {
        if (-not $SmokeGenerate) { [System.Windows.Forms.MessageBox]::Show(($problems -join "`r`n"), "请检查输入", "OK", "Warning") }
        return
    }
    $options = [pscustomobject]@{ WorksRoot = $script:worksRoot; Period = $period; OutputFile = $output; OpenAfter = $openAfterCheck.Checked }
    if (-not $SmokeGenerate) { Save-Settings ([ordered]@{ Period = $period; OutputFile = $output; OpenAfter = $openAfterCheck.Checked }) }
    $statusBox.Text = "正在读取作品图片并生成 PPT，请稍候……"
    $generateButton.Enabled = $false
    $generateButton.Text = "生成中……"
    $args = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $runScript, '-WorksRoot', $options.WorksRoot, '-Period', $period, '-OutputFile', $output)
    $argumentText = ($args | ForEach-Object { Quote-ProcessArgument "$_" }) -join ' '
    try {
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $startInfo.Arguments = $argumentText
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $startInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $startInfo.StandardErrorEncoding = [System.Text.Encoding]::UTF8
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        $null = $process.Start()
        $script:activeProcess = $process
        $script:activeOptions = $options
        $timer.Start()
    } catch {
        $generateButton.Enabled = $true
        $generateButton.Text = "生成作品展示 PPT"
        $statusBox.Text = "生成失败：`r`n$($_.Exception.Message)"
        if (-not $SmokeGenerate) { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "生成失败", "OK", "Error") }
    }
})

Refresh-Collection
if ($SmokeTest) {
    if ($form.Controls.Count -lt 15 -or -not (Test-Path -LiteralPath $runScript)) { throw "作品展示窗口初始化检查失败。" }
    Write-Output "PORTFOLIO_GUI_SMOKE_OK controls=$($form.Controls.Count)"
    $timer.Dispose(); $form.Dispose(); exit 0
}
if ($SmokeClipboard) {
    if (-not $TestImagePath -or -not (Test-Path -LiteralPath $TestImagePath)) { throw "剪贴板测试需要 -TestImagePath。" }
    $roleBox.SelectedItem = "美工"
    $nameBox.Text = "粘贴测试"
    $bitmap = [System.Drawing.Image]::FromFile($TestImagePath)
    $dataObject = New-Object System.Windows.Forms.DataObject
    $dataObject.SetData([System.Windows.Forms.DataFormats]::Bitmap, $bitmap)
    $count = Paste-ClipboardImages $dataObject
    $bitmap.Dispose()
    if ($count -ne 1 -or $collectionList.Items.Count -ne 1) { throw "剪贴板图片导入测试失败。" }
    Write-Output "PORTFOLIO_CLIPBOARD_OK images=$count people=$($collectionList.Items.Count)"
    $timer.Dispose(); $form.Dispose(); exit 0
}
if ($SmokeGenerate) {
    if (-not $TestWorksRoot -or -not (Test-Path -LiteralPath $TestWorksRoot)) { throw "完整生成测试需要 -TestWorksRoot。" }
    $script:worksRoot = $TestWorksRoot
    $script:usingClipboard = $false
    Refresh-Collection
    $outputBox.Text = Join-Path $settingsDir "作品展示完整测试.pptx"
    $openAfterCheck.Checked = $false
    $form.Add_Shown({ $generateButton.PerformClick() })
}

[void]$form.ShowDialog()
$timer.Dispose()
$form.Dispose()
if ($SmokeGenerate) {
    if ($script:smokeResult -and $script:smokeResult.Success) {
        Write-Output "PORTFOLIO_GUI_GENERATE_OK output=$($script:smokeResult.OutputFile)"
        exit 0
    }
    if ($script:smokeResult) { throw "作品展示完整生成测试失败，退出码 $($script:smokeResult.ExitCode)：$($script:smokeResult.Log)" }
    throw "作品展示完整生成测试未返回结果。"
}
