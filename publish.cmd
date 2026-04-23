@echo off
setlocal
cd /d "%~dp0"

echo === 结束运行中的程序（如果有）===
taskkill /F /IM AudioDeviceSwitcher.exe >nul 2>&1

echo === 清理旧 publish 输出 ===
if exist publish rmdir /s /q publish

echo === dotnet publish (Release, win-x64, self-contained, single-file) ===
dotnet publish CKit\CKit.csproj ^
    -c Release ^
    -r win-x64 ^
    --self-contained true ^
    -p:PublishSingleFile=true ^
    -p:EnableCompressionInSingleFile=true ^
    -o publish
if errorlevel 1 (
    echo Publish 失败
    exit /b 1
)

echo === 查找 ISCC.exe ===
set "ISCC="
if exist "%LocalAppData%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LocalAppData%\Programs\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if "%ISCC%"=="" (
    echo.
    echo 未找到 ISCC.exe，跳过安装包生成。
    echo 请确认 Inno Setup 6 已安装，或手动运行：
    echo   ISCC.exe installer\AudioDeviceSwitcher.iss
    exit /b 0
)

echo === 编译安装包: %ISCC% ===
"%ISCC%" installer\AudioDeviceSwitcher.iss
if errorlevel 1 (
    echo 安装包生成失败
    exit /b 1
)

echo.
echo === 完成 ===
echo 安装包位置：dist\
dir /b dist\*.exe 2>nul
endlocal
