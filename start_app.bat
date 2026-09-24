@echo off
title Platform Payout Reconciler
echo Starting Platform Payout Reconciler...
start http://localhost:3000
"C:\Users\user\AppData\Roaming\Antigravity\bin\agy-node.cmd" "%~dp0server.js"
pause
