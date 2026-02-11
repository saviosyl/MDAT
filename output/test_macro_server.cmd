@echo off
setlocal

REM ====== SET YOUR TOKEN HERE (or read from config later) ======
set "MACRO_TOKEN=b7f8a1d0c3e24f6a91b8d4c0a9e2f7d3c1b0a8f2e3d4c5b6a7f8e9d0c1b2a3f4"

echo Testing macro delivery...
echo.

echo 1) Health:
curl -s -i "https://mdat-macro-delivery.saviosyl.workers.dev/health"
echo.
echo.

echo 2) List:
curl -s -i -H "Authorization: Bearer %MACRO_TOKEN%" "https://mdat-macro-delivery.saviosyl.workers.dev/api/v1/list"
echo.
echo.

pause
