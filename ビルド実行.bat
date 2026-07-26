@echo off
REM ビルド実行バッチファイル
REM MZ0000_FlowchartTool rev014

echo ============================================================
echo 環境浄化とビルドを実行します
echo ============================================================

REM 環境浄化
python clean.py
if errorlevel 1 (
    echo 【エラー】環境浄化に失敗しました
    pause
    exit /b 1
)

REM ビルド実行
python build_exe.py
if errorlevel 1 (
    echo 【エラー】ビルドに失敗しました
    pause
    exit /b 1
)

echo ============================================================
echo ビルドが完了しました
echo ============================================================
pause
