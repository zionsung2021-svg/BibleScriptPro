@echo off
title BibleScript Service Launcher
chcp 65001 > nul

echo ===================================================
echo   BibleScript 서비스를 시작합니다...
echo ===================================================

:: 1. 백엔드 백그라운드 실행
start "BibleScript Backend" cmd /k "cd backend && python -m uvicorn main:app --reload --host 0.0.0.0 --port 8000"

:: 2. 프론트엔드 실행
start "BibleScript Frontend" cmd /k "cd frontend && npm run dev"

:: 3. 3초 대기 후 브라우저 자동 실행
timeout /t 3 /nobreak > nul
start http://localhost:5173

echo ===================================================
echo   실행이 완료되었습니다. 브라우저 창을 확인하세요.
echo ===================================================