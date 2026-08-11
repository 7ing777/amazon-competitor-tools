@echo off
chcp 65001 >nul
title 竞对矩阵一键分析
cd /d "%~dp0"

if "%~1"=="" (
    echo 请把【打标后的Excel文件】拖到这个窗口上, 然后按回车...
    set /p FILE=文件路径: 
) else (
    set FILE=%~1
)

echo.
echo ============================================
echo   开始分析: %FILE%
echo ============================================
echo.

python run_analysis.py --input "%FILE%"

echo.
echo ============================================
echo   完成! 输出文件已生成在同目录下
echo ============================================
pause
