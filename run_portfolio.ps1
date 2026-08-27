[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorksRoot,
    [string]$OutputFile = "",
    [string]$Period = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = [System.IO.Path]::GetFullPath((Join-Path $scriptDir ".."))
$runtimeRoot = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies"
$nodeExe = if ($env:DESIGNER_PPT_NODE) { $env:DESIGNER_PPT_NODE } else { Join-Path $runtimeRoot "node\bin\node.exe" }
$nodeModules = if ($env:DESIGNER_PPT_NODE_MODULES) { $env:DESIGNER_PPT_NODE_MODULES } else { Join-Path $runtimeRoot "node\node_modules" }
$runtimeBin = Join-Path $runtimeRoot "bin\override"

foreach ($requiredPath in @($nodeExe, $nodeModules)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) { throw "缺少运行依赖：$requiredPath" }
}
$worksResolved = (Resolve-Path -LiteralPath $WorksRoot).Path
if (-not $Period) { $Period = Get-Date -Format "yyyy年M月" }
if (-not $OutputFile) {
    $safePeriod = $Period -replace '[\\/:*?<>|]', '-'
    $OutputFile = Join-Path $projectDir "output\美工与设计师作品展示_${safePeriod}.pptx"
}
$outputResolved = [System.IO.Path]::GetFullPath($OutputFile)
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $outputResolved) | Out-Null

$buildDir = Join-Path $scriptDir ".build"
$previewDir = Join-Path $buildDir "portfolio-preview"
New-Item -ItemType Directory -Force -Path $buildDir | Out-Null
$buildResolved = [System.IO.Path]::GetFullPath($buildDir).TrimEnd([System.IO.Path]::DirectorySeparatorChar)
$previewResolved = [System.IO.Path]::GetFullPath($previewDir)
if (-not $previewResolved.StartsWith($buildResolved + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "预览目录越出构建目录，拒绝清理：$previewResolved"
}
if (Test-Path -LiteralPath $previewDir) { Remove-Item -LiteralPath $previewDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $previewDir | Out-Null

$junction = Join-Path $scriptDir "node_modules"
if (-not (Test-Path -LiteralPath $junction)) { New-Item -ItemType Junction -Path $junction -Target $nodeModules | Out-Null }

$env:RUNTIME_NODE = $nodeExe
$env:RUNTIME_NODE_MODULES = $nodeModules
$env:RUNTIME_BIN_DIR = $runtimeBin
Write-Host "[1/1] 生成美工与设计师作品展示 PPT..." -ForegroundColor Cyan
& $nodeExe (Join-Path $scriptDir "build_portfolio_deck.mjs") `
    --works-root $worksResolved `
    --output $outputResolved `
    --preview-dir $previewResolved `
    --period $Period

if ($LASTEXITCODE -ne 0) {
    $outputItem = Get-Item -LiteralPath $outputResolved -ErrorAction SilentlyContinue
    if (-not ($outputItem -and $outputItem.Length -gt 10000)) { throw "作品展示 PPT 生成失败，退出码：$LASTEXITCODE" }
    Write-Warning "渲染器在文件写完后返回退出码 $LASTEXITCODE；已确认 PPT 文件存在。"
}
$sidecar = "$outputResolved.inspect.ndjson"
if (Test-Path -LiteralPath $sidecar) { [System.IO.File]::Delete($sidecar) }
Write-Host "完成：$outputResolved" -ForegroundColor Green
Write-Host "逐页预览：$previewResolved"
