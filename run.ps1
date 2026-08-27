[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DesignerFile,

    [Parameter(Mandatory = $true)]
    [string]$SalesFile,

    [string]$ImageRoot = "",
    [string]$OutputFile = "",
    [string]$Period = "",
    [string]$DesignerSheet = "Sheet1",
    [string]$SalesSheet = "Worksheet",
    [switch]$RefreshThumbnails,
    [switch]$KeepBuildFiles
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:PYTHONIOENCODING = "utf-8"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = [System.IO.Path]::GetFullPath((Join-Path $scriptDir ".."))
$runtimeRoot = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies"
$nodeExe = if ($env:DESIGNER_PPT_NODE) { $env:DESIGNER_PPT_NODE } else { Join-Path $runtimeRoot "node\bin\node.exe" }
$pythonExe = if ($env:DESIGNER_PPT_PYTHON) { $env:DESIGNER_PPT_PYTHON } else { Join-Path $runtimeRoot "python\python.exe" }
$nodeModules = if ($env:DESIGNER_PPT_NODE_MODULES) { $env:DESIGNER_PPT_NODE_MODULES } else { Join-Path $runtimeRoot "node\node_modules" }
$runtimeBin = Join-Path $runtimeRoot "bin\override"

foreach ($requiredPath in @($nodeExe, $pythonExe, $nodeModules)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "缺少运行依赖：$requiredPath。请在 Codex Desktop 环境运行，或设置 DESIGNER_PPT_NODE / DESIGNER_PPT_PYTHON / DESIGNER_PPT_NODE_MODULES。"
    }
}

$designerResolved = (Resolve-Path -LiteralPath $DesignerFile).Path
$salesResolved = (Resolve-Path -LiteralPath $SalesFile).Path
if (-not $ImageRoot) {
    $ImageRoot = Join-Path $projectDir "work\designer_sales_2026-08\assets"
}
$imageResolved = (Resolve-Path -LiteralPath $ImageRoot).Path

if (-not $Period) {
    $match = [regex]::Match([System.IO.Path]::GetFileName($salesResolved), '(20\d{2})[-_.年](0?[1-9]|1[0-2])')
    $Period = if ($match.Success) { "$($match.Groups[1].Value)年$([int]$match.Groups[2].Value)月" } else { Get-Date -Format "yyyy年M月" }
}
if (-not $OutputFile) {
    $safePeriod = $Period -replace '[\\/:*?<>|]', '-'
    $OutputFile = Join-Path $projectDir "output\设计师型号交易净利_${safePeriod}_可编辑版.pptx"
}
$outputResolved = [System.IO.Path]::GetFullPath($OutputFile)
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $outputResolved) | Out-Null

$buildDir = Join-Path $scriptDir ".build"
$previewDir = Join-Path $buildDir "preview"
New-Item -ItemType Directory -Force -Path $buildDir | Out-Null
$buildResolved = [System.IO.Path]::GetFullPath($buildDir).TrimEnd([System.IO.Path]::DirectorySeparatorChar)
$previewResolved = [System.IO.Path]::GetFullPath($previewDir)
if (-not $previewResolved.StartsWith($buildResolved + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "预览目录越出构建目录，拒绝清理：$previewResolved"
}
if (Test-Path -LiteralPath $previewDir) {
    Remove-Item -LiteralPath $previewDir -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $previewDir | Out-Null

$junction = Join-Path $scriptDir "node_modules"
if (-not (Test-Path -LiteralPath $junction)) {
    New-Item -ItemType Junction -Path $junction -Target $nodeModules | Out-Null
}

$dataJson = Join-Path $buildDir "designer-sales.json"
$cropConfig = Join-Path $scriptDir "thumbnail_crops.json"
$thumbnailArgs = @(
    (Join-Path $scriptDir "make_thumbnails.py"),
    "--image-root", $imageResolved,
    "--config", $cropConfig
)
if ($RefreshThumbnails) { $thumbnailArgs += "--refresh" }

Write-Host "[1/3] 检查型号缩略图..." -ForegroundColor Cyan
& $pythonExe @thumbnailArgs
if ($LASTEXITCODE -ne 0) { throw "缩略图处理失败，退出码：$LASTEXITCODE" }

Write-Host "[2/3] 汇总设计师型号交易额与净利..." -ForegroundColor Cyan
& $pythonExe (Join-Path $scriptDir "extract_data.py") `
    --designer-file $designerResolved `
    --sales-file $salesResolved `
    --output $dataJson `
    --designer-sheet $DesignerSheet `
    --sales-sheet $SalesSheet `
    --period $Period
if ($LASTEXITCODE -ne 0) { throw "数据提取失败，退出码：$LASTEXITCODE" }

$env:RUNTIME_NODE = $nodeExe
$env:RUNTIME_NODE_MODULES = $nodeModules
$env:RUNTIME_BIN_DIR = $runtimeBin
Write-Host "[3/3] 生成可编辑的月度业绩 PPT..." -ForegroundColor Cyan
& $nodeExe (Join-Path $scriptDir "build_deck.mjs") `
    --input $dataJson `
    --output $outputResolved `
    --image-root $imageResolved `
    --preview-dir $previewDir `
    --period $Period `
    --crop-config $cropConfig

if ($LASTEXITCODE -ne 0) {
    $outputItem = Get-Item -LiteralPath $outputResolved -ErrorAction SilentlyContinue
    if ($outputItem -and $outputItem.Length -gt 10000) {
        Write-Warning "渲染器在文件写完后返回退出码 $LASTEXITCODE；已确认 PPT 文件存在，继续完成。"
    }
    else {
        throw "PPT 生成失败，退出码：$LASTEXITCODE"
    }
}

if (-not $KeepBuildFiles) {
    Remove-Item -LiteralPath $dataJson -Force -ErrorAction SilentlyContinue
}

Write-Host "完成：$outputResolved" -ForegroundColor Green
Write-Host "逐页预览：$previewDir"
