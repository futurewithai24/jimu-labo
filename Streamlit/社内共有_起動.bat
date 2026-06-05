@echo off
chcp 65001 > nul
cd /d E:\Antigravity\Streamlit

echo ====================================
echo  じむラボ 社内共有モードで起動中...
echo ====================================
echo.

:: ファイアウォールにポート8501を許可（初回のみ有効・管理者権限で実行時）
netsh advfirewall firewall show rule name="Streamlit 8501" > nul 2>&1
if errorlevel 1 (
    echo ファイアウォール設定を追加しています...
    netsh advfirewall firewall add rule name="Streamlit 8501" dir=in action=allow protocol=TCP localport=8501 > nul 2>&1
    echo 完了しました。
    echo.
)

:: IPアドレスを自動取得
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    set IP=%%a
    goto :found
)
:found
set IP=%IP: =%

echo 【社内の人に教えるURL】
echo.
echo   http://%IP%:8501
echo.
echo ※ このウィンドウを閉じるとアプリが止まります
echo ====================================
echo.

streamlit run app.py --server.address 0.0.0.0 --server.port 8501
pause
