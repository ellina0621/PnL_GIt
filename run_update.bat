@echo off
chcp 65001 >nul
cd /d "D:\我才不要走量化\PnL_GIt"

echo [%date% %time%] ===== 開始每日更新 =====

python rebuild_daily_pnl.py
if %errorlevel% neq 0 (
    echo [錯誤] rebuild_daily_pnl.py 執行失敗，中止。
    exit /b 1
)

python export_json.py
if %errorlevel% neq 0 (
    echo [錯誤] export_json.py 執行失敗，中止。
    exit /b 1
)

git add "trading-journal (2) - 每日損益重算.xlsx" "trading-journal (2).xlsx" docs/data.json
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo [資訊] 無變動，略過 commit。
) else (
    git commit -m "auto update %date%"
    git push
    echo [完成] 已推送至 GitHub Pages。
)

echo [%date% %time%] ===== 更新完成 =====
