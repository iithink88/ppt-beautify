# export_preview.ps1
# 用 PowerPoint / WPS 演示 COM 把每页导出为 JPG，用于肉眼视觉 QA。
# 用法: powershell -File export_preview.ps1 输入.pptx [输出目录]
# 本机需安装 PowerPoint 或 WPS 演示（无需管理员）。
param($pptx, $outDir)

if (-not $pptx) { Write-Host "usage: export_preview.ps1 输入.pptx [输出目录]"; exit 1 }

$out = if ($outDir) { $outDir } else { "$env:TEMP\ppt_preview" }
if (Test-Path $out) { Remove-Item $out -Recurse -Force }
New-Item -ItemType Directory -Path $out | Out-Null

# 优先 PowerPoint，回退 WPS 演示
$progId = $null
foreach ($p in @("PowerPoint.Application", "kwpp.Application")) {
    try {
        $t = [Type]::GetTypeFromProgID($p)
        if ($t) { $progId = $p; break }
    } catch { }
}
if (-not $progId) { Write-Host "未找到 PowerPoint 或 WPS 演示 COM 组件"; exit 1 }

$pp = New-Object -ComObject $progId
try {
    $pres = $pp.Presentations.Open($pptx, $true, $false, $false)
    $i = 0
    foreach ($s in $pres.Slides) {
        $i++
        $f = Join-Path $out ("slide{0:D2}.jpg" -f $i)
        $s.Export($f, "JPG", 1600, 900)
    }
    $pres.Close()
    Write-Host "EXPORTED $i slides -> $out (via $progId)"
} finally {
    $pp.Quit()
}
