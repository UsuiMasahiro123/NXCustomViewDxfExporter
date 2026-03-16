# NXCustomViewDxfExporter Build & Sign Script
$ErrorActionPreference = "Stop"

$ProjectDir = "C:\Users\tcadmin\source\repos\NXCustomViewDxfExporter\NXCustomViewDxfExporter"
$CsFile = "$ProjectDir\CustomViewDxfExporter.cs"
$CsprojFile = "$ProjectDir\NXCustomViewDxfExporter.csproj"
$DllFile = "$ProjectDir\bin\x64\Debug\NXCustomViewDxfExporter.dll"
$MSBuild = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
$SignTool = "C:\Program Files\Siemens\NX2312\NXBIN\SignDotNet.exe"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " NXCustomViewDxfExporter Build & Sign" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: CSファイルの最終更新日時
$csInfo = Get-Item $CsFile
Write-Host "[Step 1] CS file: $($csInfo.LastWriteTime)  ($([math]::Round($csInfo.Length/1024,1)) KB)" -ForegroundColor Yellow

# Step 2: NXプロセス終了 + クリーンビルド
Write-Host "[Step 2] Stopping NX process..." -ForegroundColor Yellow
Stop-Process -Name ugraf -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

Write-Host "[Step 3] Clean + Build (Debug/x64)..." -ForegroundColor Yellow
& "$MSBuild" /t:Rebuild /p:Configuration=Debug /p:Platform=x64 /v:minimal "$CsprojFile"

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "[ERROR] Build FAILED! (exit code: $LASTEXITCODE)" -ForegroundColor Red
    exit 1
}

# Step 4: DLL更新日時確認
if (!(Test-Path $DllFile)) {
    Write-Host "[ERROR] DLL not found: $DllFile" -ForegroundColor Red
    exit 1
}

$dllInfo = Get-Item $DllFile
Write-Host ""
Write-Host "[Step 4] DLL file: $($dllInfo.LastWriteTime)  ($([math]::Round($dllInfo.Length/1024,1)) KB)" -ForegroundColor Yellow

if ($dllInfo.LastWriteTime -lt $csInfo.LastWriteTime) {
    Write-Host "[WARNING] DLL is OLDER than CS file!" -ForegroundColor Red
} else {
    Write-Host "  -> DLL is newer than CS file (OK)" -ForegroundColor Green
}

# Step 5: 署名
Write-Host "[Step 5] Signing with SignDotNet.exe..." -ForegroundColor Yellow
& "$SignTool" "$DllFile"

if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Signing FAILED! (exit code: $LASTEXITCODE)" -ForegroundColor Red
    exit 1
}

# Step 6: 完了
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " BUILD & SIGN COMPLETE" -ForegroundColor Green
Write-Host " DLL: $DllFile" -ForegroundColor Green
Write-Host " Time: $($dllInfo.LastWriteTime)" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
