param (
    [string]$selectedMatch = $env:ONE_CLICK_RUN_PORTABLE_PYTHON_SELECTEDMATCH
)

$ErrorActionPreference = "Stop"

# 前缀
$固定前缀 = "OCR-portable-python"

# 获取可选版本
if (-not $selectedMatch) {
    Write-Output "获取可选择的版本..."
    $url = "https://github.com/one-click-run/portable-python/releases/expanded_assets/python"
    $response = Invoke-WebRequest -Uri $url
    $htmlContent = $response.Content
    $pattern = 'python-(\d+\.\d+\.\d+)-amd64.zip'
    $matches = [regex]::Matches($htmlContent, $pattern)
    $uniqueMatches = $matches | ForEach-Object { $_.Value } | Sort-Object -Unique

    # 弹出选择界面
    $selectedMatch = $uniqueMatches | Out-GridView -Title "Select a match" -OutputMode Single

    # 如果用户没有选择任何版本，退出脚本
    if ([string]::IsNullOrEmpty($selectedMatch)) {
        Write-Output "用户没有选择任何版本, 脚本将退出..."
        exit
    }
}

Write-Output "用户选择了: $selectedMatch"

# 弹出输入框让用户自定义安装名称
Add-Type -AssemblyName Microsoft.VisualBasic
$默认名称 = "default"
$用户输入名称 = [Microsoft.VisualBasic.Interaction]::InputBox("请输入安装名称:", "安装路径", $默认名称)
if ([string]::IsNullOrWhiteSpace($用户输入名称)) {
    $用户输入名称 = $默认名称
}
Write-Output "安装名称为: $用户输入名称"

# 构建文件夹和文件名
$安装主文件夹 = "$固定前缀-$用户输入名称"
$虚拟环境文件夹 = "$安装主文件夹-venv"
$修复脚本文件 = "$安装主文件夹-fix.cmd"
$启动脚本文件 = "$安装主文件夹.cmd"

# 下载
Write-Output "开始下载..."
$downloadUrl = "https://github.com/one-click-run/portable-python/releases/download/python/$selectedMatch"
$localFileName = "$安装主文件夹.zip"
Invoke-WebRequest -Uri $downloadUrl -OutFile $localFileName
Write-Output "下载完成..."

# 解压
Write-Output "开始解压..."
$extractedFolder = [System.IO.Path]::GetFileNameWithoutExtension($localFileName)
Expand-Archive -Path $localFileName -DestinationPath $安装主文件夹 -Force
Write-Output "解压完成..."

# 创建虚拟环境
Write-Output "开始创建虚拟环境..."
Remove-Item -Path ".\$虚拟环境文件夹" -Recurse -Force -ErrorAction SilentlyContinue
$venvCommand = Join-Path $安装主文件夹 "python.exe"
& $venvCommand -m venv $虚拟环境文件夹
Write-Output "虚拟环境创建完成..."

# 删除压缩包
Remove-Item -Path $localFileName -Force -ErrorAction SilentlyContinue

# 替换 activate 文件
$scriptPath = $PWD.Path
$activateFilePath = Join-Path $虚拟环境文件夹 "Scripts\activate"
$activateContent = Get-Content -Path $activateFilePath -Raw
$activateContent = $activateContent -replace [regex]::Escape("$scriptPath\venv"), '$( cd $( dirname ${BASH_SOURCE[0]} ) && pwd )/../'
[System.IO.File]::WriteAllLines($activateFilePath, $activateContent)

# 替换 activate.bat 文件
$activateFilePath = Join-Path $虚拟环境文件夹 "Scripts\activate.bat"
$activateContent = Get-Content -Path $activateFilePath -Raw
$activateContent = $activateContent -replace [regex]::Escape("$scriptPath\venv"), '%~dp0\..'
[System.IO.File]::WriteAllLines($activateFilePath, $activateContent)

# 创建修复脚本
$scriptContent = @"
setlocal enabledelayedexpansion

set "charset=abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
set "length=8"
set "randStr="
for /l %%i in (1,1,%length%) do (
    set /a "index=!random! %% 62"
    for %%j in (!index!) do set "randStr=!randStr!!charset:~%%j,1!"
)

set "venvTemp=$虚拟环境文件夹-!randStr!"
rename $虚拟环境文件夹 !venvTemp!

set venvCommand=.\$安装主文件夹\python.exe
%venvCommand% -m venv $虚拟环境文件夹

for /d %%d in (.\"!venvTemp!\"\*) do (
    if /i not "%%~nxd"=="Scripts" (
        if exist .\$虚拟环境文件夹\"%%~nxd" rmdir /s /q .\$虚拟环境文件夹\"%%~nxd"
        move "%%d" .\$虚拟环境文件夹
    )
)

rmdir /s /q !venvTemp!

endlocal

start .\$虚拟环境文件夹\Scripts\activate.bat
"@
$scriptFilePath = Join-Path $scriptPath $修复脚本文件
[System.IO.File]::WriteAllLines($scriptFilePath, $scriptContent)

# 创建启动脚本
$scriptContent = @"
@echo off
start .\$虚拟环境文件夹\Scripts\activate.bat
"@
$scriptFilePath = Join-Path $scriptPath $启动脚本文件
[System.IO.File]::WriteAllLines($scriptFilePath, $scriptContent)

Write-Output "完成"
