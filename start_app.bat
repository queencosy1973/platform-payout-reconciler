@echo off
chcp 65001 > nul
title Platform Payout Reconciler Launcher
echo =====================================================================
echo    ระบบตรวจสอบและกระทบยอดเงินเข้า (Platform Payout Reconciler)
echo =====================================================================
echo.
echo กำลังเปิดระบบในเบราว์เซอร์...
start "" "%~dp0index.html"
echo.
echo ระบบพร้อมใช้งานแล้วบนเบราว์เซอร์!
echo.
pause
