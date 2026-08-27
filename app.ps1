param(
    [switch]$SmokeTest,
    [switch]$SmokeGenerate,
    [string]$TestWorksRoot = ""
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = [System.IO.Path]::GetFullPath((Join-Path $scriptDir ".."))
$runScript = Join-Path $scriptDir "run_monthly_meeting.ps1"
$portfolioApp = Join-Path $scriptDir "portfolio_app.ps1"
$templateFile = Join-Path $scriptDir "型号归属表模板.xlsx"
$settingsDir = Join-Path $scriptDir ".build"
$settingsFile = Join-Path $settingsDir "gui-settings.json"
$portfolioSourceFile = Join-Path $settingsDir "portfolio-source.txt"
$clipboardWorksRoot = Join-Path $settingsDir "clipboard-portfolio"
New-Item -ItemType Directory -Force -Path $settingsDir | Out-Null

function Infer-Period([string]$filePath) {
    $name = [System.IO.Path]::GetFileName($filePath)
    $match = [regex]::Match($name, '(20\d{2})[-_.年](0?[1-9]|1[0-2])')
    if ($match.Success) {
        return "$($match.Groups[1].Value)年$([int]$match.Groups[2].Value)月"
    }
    return (Get-Date -Format "yyyy年M月")
}

function Default-Output([string]$period) {
    $safePeriod = $period -replace '[\\/:*?<>|]', '-'
    return (Join-Path $projectDir "output\美工月会_${safePeriod}_完整可编辑版.pptx")
}

function Quote-ProcessArgument([string]$value) {
    return '"' + $value.Replace('"', '\"') + '"'
}

function Save-Settings($values) {
    $values | ConvertTo-Json | Set-Content -LiteralPath $settingsFile -Encoding UTF8
}

$defaultDesigner = if (Test-Path -LiteralPath 'D:\desktop\分红型号.xlsx') {
    'D:\desktop\分红型号.xlsx'
} else {
    $templateFile
}
$defaultSales = ""
$downloads = Join-Path $env:USERPROFILE "Downloads"
if (Test-Path -LiteralPath $downloads) {
    $latest = Get-ChildItem -LiteralPath $downloads -Filter 'amt_detail_*.xlsx' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($latest) { $defaultSales = $latest.FullName }
}
$defaultImages = Join-Path $projectDir "work\designer_sales_2026-08\assets"
$defaultPeriod = Infer-Period $defaultSales
$defaults = [ordered]@{
    DesignerFile = $defaultDesigner
    SalesFile = $defaultSales
    ImageRoot = $defaultImages
    Period = $defaultPeriod
    OutputFile = Default-Output $defaultPeriod
    OpenAfter = $true
}
if (Test-Path -LiteralPath $settingsFile) {
    try {
        $saved = Get-Content -Raw -LiteralPath $settingsFile | ConvertFrom-Json
        foreach ($name in @('DesignerFile', 'SalesFile', 'ImageRoot', 'Period', 'OutputFile', 'OpenAfter')) {
            if ($null -ne $saved.$name -and "$($saved.$name)" -ne "") {
                $defaults[$name] = $saved.$name
            }
        }
    } catch {
        # 设置损坏时回退到默认值，不影响主流程。
    }
}
if ([System.IO.Path]::GetFileName("$($defaults.OutputFile)").StartsWith("设计师型号交易净利_")) {
    $defaults.OutputFile = Default-Output "$($defaults.Period)"
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "设计师型号月报助手"
$appIcon = Join-Path $scriptDir "assets\月报助手.ico"
if (Test-Path -LiteralPath $appIcon) { $form.Icon = New-Object System.Drawing.Icon -ArgumentList $appIcon }
$form.StartPosition = "CenterScreen"
$form.ClientSize = New-Object System.Drawing.Size(850, 700)
$form.MinimumSize = New-Object System.Drawing.Size(866, 739)
$form.BackColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)

$header = New-Object System.Windows.Forms.Panel
$header.Location = New-Object System.Drawing.Point(0, 0)
$header.Size = New-Object System.Drawing.Size(850, 82)
$header.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#243F8E")
$header.Anchor = "Top,Left,Right"
$form.Controls.Add($header)

$title = New-Object System.Windows.Forms.Label
$title.Text = "设计师型号月报助手"
$title.ForeColor = [System.Drawing.Color]::White
$title.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 18, [System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(26, 14)
$header.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "一份 PPT 同时展示作品与型号月度销售额、利润"
$subtitle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#DCE6FF")
$subtitle.AutoSize = $true
$subtitle.Location = New-Object System.Drawing.Point(29, 51)
$header.Controls.Add($subtitle)

$guide = New-Object System.Windows.Forms.Label
$guide.Text = "先点橙色按钮录入作品，再选择交易表；蓝色按钮生成一份完整月会 PPT。"
$guide.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#24489E")
$guide.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10, [System.Drawing.FontStyle]::Bold)
$guide.AutoSize = $true
$guide.Location = New-Object System.Drawing.Point(30, 98)
$form.Controls.Add($guide)

function Add-FileRow([int]$top, [string]$labelText, [string]$initial, [string]$buttonText) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Location = New-Object System.Drawing.Point -ArgumentList 30, ($top + 7)
    $label.Size = New-Object System.Drawing.Size(145, 28)
    $form.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Text = $initial
    $textbox.Location = New-Object System.Drawing.Point -ArgumentList 180, $top
    $textbox.Size = New-Object System.Drawing.Size(545, 30)
    $textbox.Anchor = "Top,Left,Right"
    $form.Controls.Add($textbox)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $buttonText
    $button.Location = New-Object System.Drawing.Point -ArgumentList 738, ($top - 1)
    $button.Size = New-Object System.Drawing.Size(82, 32)
    $button.Anchor = "Top,Right"
    $form.Controls.Add($button)
    return @($textbox, $button)
}

$designerRow = Add-FileRow 135 "1. 型号归属表" $defaults.DesignerFile "选择..."
$designerBox = $designerRow[0]
$designerButton = $designerRow[1]
$salesRow = Add-FileRow 185 "2. 当月交易明细" $defaults.SalesFile "选择..."
$salesBox = $salesRow[0]
$salesButton = $salesRow[1]
$imageRow = Add-FileRow 235 "3. 型号图片文件夹" $defaults.ImageRoot "选择..."
$imageBox = $imageRow[0]
$imageButton = $imageRow[1]

$periodLabel = New-Object System.Windows.Forms.Label
$periodLabel.Text = "4. 报告月份"
$periodLabel.Location = New-Object System.Drawing.Point(30, 292)
$periodLabel.Size = New-Object System.Drawing.Size(145, 28)
$form.Controls.Add($periodLabel)
$periodBox = New-Object System.Windows.Forms.TextBox
$periodBox.Text = $defaults.Period
$periodBox.Location = New-Object System.Drawing.Point(180, 285)
$periodBox.Size = New-Object System.Drawing.Size(180, 30)
$form.Controls.Add($periodBox)
$periodHint = New-Object System.Windows.Forms.Label
$periodHint.Text = "例如：2026年8月"
$periodHint.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#657086")
$periodHint.Location = New-Object System.Drawing.Point(375, 292)
$periodHint.AutoSize = $true
$form.Controls.Add($periodHint)

$outputRow = Add-FileRow 335 "5. 完整 PPT 保存位置" $defaults.OutputFile "保存为..."
$outputBox = $outputRow[0]
$outputButton = $outputRow[1]

$templateButton = New-Object System.Windows.Forms.Button
$templateButton.Text = "打开型号归属表模板"
$templateButton.Location = New-Object System.Drawing.Point(180, 382)
$templateButton.Size = New-Object System.Drawing.Size(180, 36)
$templateButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#EEF4FF")
$templateButton.FlatStyle = "Flat"
$templateButton.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml("#B8C3D6")
$form.Controls.Add($templateButton)

$openImagesButton = New-Object System.Windows.Forms.Button
$openImagesButton.Text = "打开型号图片文件夹"
$openImagesButton.Location = New-Object System.Drawing.Point(375, 382)
$openImagesButton.Size = New-Object System.Drawing.Size(180, 36)
$openImagesButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F5F7FA")
$openImagesButton.FlatStyle = "Flat"
$openImagesButton.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml("#B8C3D6")
$form.Controls.Add($openImagesButton)

$refreshCheck = New-Object System.Windows.Forms.CheckBox
$refreshCheck.Text = "原图有更新时，重新生成缩略图"
$refreshCheck.Location = New-Object System.Drawing.Point(575, 388)
$refreshCheck.Size = New-Object System.Drawing.Size(245, 26)
$form.Controls.Add($refreshCheck)

$modelHelp = New-Object System.Windows.Forms.Label
$modelHelp.Text = "型号放法：在《型号归属表》Sheet1 中填写 型号、设计师；图片按 图片文件夹\设计师\基础型号.png 放置。"
$modelHelp.Location = New-Object System.Drawing.Point(30, 435)
$modelHelp.Size = New-Object System.Drawing.Size(790, 28)
$modelHelp.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#657086")
$form.Controls.Add($modelHelp)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.ReadOnly = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.Location = New-Object System.Drawing.Point(30, 470)
$statusBox.Size = New-Object System.Drawing.Size(790, 120)
$statusBox.Anchor = "Top,Bottom,Left,Right"
$statusBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F7F9FF")
$statusBox.Text = "等待生成。PPT 页序固定：①美工/设计师作品 ②型号销售额与利润。`r`n交易表需包含：型号、有效销售、净利。"
$form.Controls.Add($statusBox)

$openAfterCheck = New-Object System.Windows.Forms.CheckBox
$openAfterCheck.Text = "完成后自动打开 PPT"
$openAfterCheck.Checked = [bool]$defaults.OpenAfter
$openAfterCheck.Location = New-Object System.Drawing.Point(30, 615)
$openAfterCheck.Size = New-Object System.Drawing.Size(190, 28)
$openAfterCheck.Anchor = "Bottom,Left"
$form.Controls.Add($openAfterCheck)

$portfolioButton = New-Object System.Windows.Forms.Button
$portfolioButton.Text = "录入美工/设计师作品"
$portfolioButton.Location = New-Object System.Drawing.Point(330, 608)
$portfolioButton.Size = New-Object System.Drawing.Size(235, 48)
$portfolioButton.Anchor = "Bottom,Right"
$portfolioButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F26B38")
$portfolioButton.ForeColor = [System.Drawing.Color]::White
$portfolioButton.FlatStyle = "Flat"
$portfolioButton.FlatAppearance.BorderSize = 0
$portfolioButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($portfolioButton)

$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = "生成完整月会 PPT"
$generateButton.Location = New-Object System.Drawing.Point(585, 608)
$generateButton.Size = New-Object System.Drawing.Size(235, 48)
$generateButton.Anchor = "Bottom,Right"
$generateButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#336CFF")
$generateButton.ForeColor = [System.Drawing.Color]::White
$generateButton.FlatStyle = "Flat"
$generateButton.FlatAppearance.BorderSize = 0
$generateButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($generateButton)

$openExcelDialog = New-Object System.Windows.Forms.OpenFileDialog
$openExcelDialog.Filter = "Excel 文件 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|所有文件 (*.*)|*.*"
$openExcelDialog.CheckFileExists = $true
$savePptDialog = New-Object System.Windows.Forms.SaveFileDialog
$savePptDialog.Filter = "PowerPoint 文件 (*.pptx)|*.pptx"
$savePptDialog.AddExtension = $true
$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog

$designerButton.Add_Click({
    $openExcelDialog.Title = "选择型号归属表"
    if ($designerBox.Text -and (Test-Path -LiteralPath $designerBox.Text)) { $openExcelDialog.InitialDirectory = Split-Path -Parent $designerBox.Text }
    if ($openExcelDialog.ShowDialog() -eq "OK") { $designerBox.Text = $openExcelDialog.FileName }
})
$salesButton.Add_Click({
    $openExcelDialog.Title = "选择当月交易明细"
    if ($salesBox.Text -and (Test-Path -LiteralPath $salesBox.Text)) { $openExcelDialog.InitialDirectory = Split-Path -Parent $salesBox.Text }
    if ($openExcelDialog.ShowDialog() -eq "OK") {
        $salesBox.Text = $openExcelDialog.FileName
        $periodBox.Text = Infer-Period $salesBox.Text
        $outputBox.Text = Default-Output $periodBox.Text
    }
})
$imageButton.Add_Click({
    $folderDialog.Description = "选择按设计师分类的型号图片根目录"
    if ($imageBox.Text -and (Test-Path -LiteralPath $imageBox.Text)) { $folderDialog.SelectedPath = $imageBox.Text }
    if ($folderDialog.ShowDialog() -eq "OK") { $imageBox.Text = $folderDialog.SelectedPath }
})
$outputButton.Add_Click({
    $savePptDialog.FileName = [System.IO.Path]::GetFileName($outputBox.Text)
    if ($outputBox.Text) { $savePptDialog.InitialDirectory = Split-Path -Parent $outputBox.Text }
    if ($savePptDialog.ShowDialog() -eq "OK") { $outputBox.Text = $savePptDialog.FileName }
})
$templateButton.Add_Click({
    if (Test-Path -LiteralPath $templateFile) { Start-Process -FilePath $templateFile } else { [System.Windows.Forms.MessageBox]::Show("模板文件不存在：$templateFile", "提示") }
})
$openImagesButton.Add_Click({
    if (Test-Path -LiteralPath $imageBox.Text) { Start-Process -FilePath "explorer.exe" -ArgumentList (Quote-ProcessArgument $imageBox.Text) } else { [System.Windows.Forms.MessageBox]::Show("图片文件夹不存在。", "提示") }
})
$periodBox.Add_Leave({
    if ($periodBox.Text.Trim()) { $outputBox.Text = Default-Output $periodBox.Text.Trim() }
})
$portfolioButton.Add_Click({
    if (-not (Test-Path -LiteralPath $portfolioApp)) {
        [System.Windows.Forms.MessageBox]::Show("作品展示功能文件不存在：$portfolioApp", "提示", "OK", "Error")
        return
    }
    Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -ArgumentList @('-NoProfile', '-WindowStyle', 'Hidden', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $portfolioApp + '"')) `
        -WindowStyle Hidden
})

$script:activeProcess = $null
$script:activeOptions = $null
$script:smokeResult = $null
$generationTimer = New-Object System.Windows.Forms.Timer
$generationTimer.Interval = 500
$generationTimer.Add_Tick({
    if (-not $script:activeProcess) {
        $generationTimer.Stop()
        return
    }
    if (-not $script:activeProcess.HasExited) { return }

    $generationTimer.Stop()
    $process = $script:activeProcess
    $options = $script:activeOptions
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $exitCode = $process.ExitCode
    $process.Dispose()
    $script:activeProcess = $null
    $script:activeOptions = $null

    $generateButton.Enabled = $true
    $generateButton.Text = "生成完整月会 PPT"
    $statusBox.Text = ($stdout + "`r`n" + $stderr).Trim()
    if ($SmokeGenerate) {
        $script:smokeResult = [pscustomobject]@{
            Success = ($exitCode -eq 0 -and (Test-Path -LiteralPath $options.OutputFile))
            ExitCode = $exitCode
            OutputFile = $options.OutputFile
            Log = $statusBox.Text
        }
        $form.Close()
        return
    }
    if ($exitCode -eq 0 -and (Test-Path -LiteralPath $options.OutputFile)) {
        [System.Windows.Forms.MessageBox]::Show("完整月会 PPT 已生成：`r`n$($options.OutputFile)", "完成", "OK", "Information")
        if ($options.OpenAfter) { Start-Process -FilePath $options.OutputFile }
    } else {
        [System.Windows.Forms.MessageBox]::Show("生成失败，请查看界面中的运行日志。", "生成失败", "OK", "Error")
    }
})

$generateButton.Add_Click({
    $designer = $designerBox.Text.Trim()
    $sales = $salesBox.Text.Trim()
    $images = $imageBox.Text.Trim()
    $period = $periodBox.Text.Trim()
    $output = $outputBox.Text.Trim()
    $works = if ($TestWorksRoot) {
        $TestWorksRoot
    } elseif (Test-Path -LiteralPath $portfolioSourceFile -PathType Leaf) {
        (Get-Content -Raw -LiteralPath $portfolioSourceFile).Trim()
    } else {
        $clipboardWorksRoot
    }
    $problems = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $designer -PathType Leaf)) { $problems.Add("请选择有效的型号归属表。") }
    if (-not (Test-Path -LiteralPath $sales -PathType Leaf)) { $problems.Add("请选择有效的当月交易明细。") }
    if (-not (Test-Path -LiteralPath $images -PathType Container)) { $problems.Add("请选择有效的型号图片文件夹。") }
    $worksImages = if (Test-Path -LiteralPath $works -PathType Container) {
        @(Get-ChildItem -LiteralPath $works -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension.ToLowerInvariant() -in @('.png', '.jpg', '.jpeg', '.webp') })
    } else { @() }
    if ($worksImages.Count -eq 0) { $problems.Add("请先点击橙色按钮，录入美工或设计师作品图片。") }
    if (-not $period) { $problems.Add("请填写报告月份。") }
    if (-not $output -or [System.IO.Path]::GetExtension($output).ToLowerInvariant() -ne ".pptx") { $problems.Add("请选择 .pptx 输出文件。") }
    if ($problems.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(($problems -join "`r`n"), "请检查输入", "OK", "Warning")
        return
    }
    $settings = [ordered]@{
        DesignerFile = $designer
        SalesFile = $sales
        WorksRoot = $works
        ImageRoot = $images
        Period = $period
        OutputFile = $output
        OpenAfter = $openAfterCheck.Checked
    }
    if (-not $SmokeGenerate) { Save-Settings $settings }
    $statusBox.Text = "正在生成一份完整月会 PPT，请稍候……"
    $generateButton.Enabled = $false
    $generateButton.Text = "生成中……"
    $options = [pscustomobject]@{
        DesignerFile = $designer
        SalesFile = $sales
        WorksRoot = $works
        ImageRoot = $images
        Period = $period
        OutputFile = $output
        RefreshThumbnails = $refreshCheck.Checked
        OpenAfter = $openAfterCheck.Checked
    }
    $processArgs = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $runScript,
        '-DesignerFile', $options.DesignerFile,
        '-SalesFile', $options.SalesFile,
        '-WorksRoot', $options.WorksRoot,
        '-ImageRoot', $options.ImageRoot,
        '-Period', $options.Period,
        '-OutputFile', $options.OutputFile
    )
    if ($options.RefreshThumbnails) { $processArgs += '-RefreshThumbnails' }
    $argumentText = ($processArgs | ForEach-Object { Quote-ProcessArgument "$_" }) -join ' '
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
        $startInfo.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8"
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        $null = $process.Start()
        $script:activeProcess = $process
        $script:activeOptions = $options
        $generationTimer.Start()
    } catch {
        $generateButton.Enabled = $true
        $generateButton.Text = "生成完整月会 PPT"
        $statusBox.Text = "生成失败：`r`n$($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "生成失败", "OK", "Error")
    }
})

if ($SmokeTest) {
    if ($form.Controls.Count -lt 10 -or -not (Test-Path -LiteralPath $runScript)) { throw "GUI 初始化检查失败。" }
    Write-Output "GUI_SMOKE_OK controls=$($form.Controls.Count)"
    $generationTimer.Dispose()
    $form.Dispose()
    exit 0
}

if ($SmokeGenerate) {
    foreach ($requiredPath in @($designerBox.Text, $salesBox.Text, $imageBox.Text, $TestWorksRoot)) {
        if (-not (Test-Path -LiteralPath $requiredPath)) { throw "完整生成测试缺少输入：$requiredPath" }
    }
    $outputBox.Text = Join-Path $settingsDir "GUI完整生成测试.pptx"
    $openAfterCheck.Checked = $false
    $form.Add_Shown({ $generateButton.PerformClick() })
}

[void]$form.ShowDialog()
$generationTimer.Dispose()
$form.Dispose()

if ($SmokeGenerate) {
    if ($script:smokeResult -and $script:smokeResult.Success) {
        Write-Output "GUI_GENERATE_OK output=$($script:smokeResult.OutputFile)"
        exit 0
    }
    if ($script:smokeResult) {
        throw "GUI 完整生成测试失败，退出码 $($script:smokeResult.ExitCode)：$($script:smokeResult.Log)"
    }
    throw "GUI 完整生成测试未返回结果。"
}
