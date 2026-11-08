@echo off
echo ====================================================================
echo   MYLIFE: OXFORD 5000 VOCABULARY ^& SMART FINANCE TRACKER - APK BUILD
echo ====================================================================
echo.
echo Memeriksa instalasi Flet...
pip install "flet[all]" openpyxl
echo.
echo Memulai proses kompilasi Python ke Android APK...
echo Catatan: Pastikan Android SDK sudah terinstal untuk build lokal,
echo atau gunakan GitHub Actions (.github/workflows/build_apk.yml) untuk build otomatis di cloud.
echo.
flet build apk --verbose
pause

