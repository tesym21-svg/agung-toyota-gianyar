@echo off
title AGUNG TOYOTA GIANYAR - Sistem Pengingat Otomatis WA & Email (Design by MDP)
cd /d "%~dp0"
echo ============================================================================
echo   AGUNG TOYOTA GIANYAR - SMART REMINDER SYSTEM (Design by MDP)
echo   Jam Operasional WA Kantor : 08:00 - 16:00 WITA
echo   Jadwal Pengiriman Otomatis: 10:00 - 14:00 WITA
echo ============================================================================
echo.
echo Membuka Dashboard di Browser (http://localhost:3000)...
start http://localhost:3000
node server.js
pause
