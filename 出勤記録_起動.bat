@echo off
cd /d "%~dp0"
echo 出勤記録の自動集計を開始します...
python kintai_mail_collector.py
if %errorlevel% neq 0 (
    echo.
    echo エラーが発生しました。Pythonがインストールされているか確認してください。
)
pause
