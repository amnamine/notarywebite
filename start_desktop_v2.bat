@echo off
chcp 65001 > nul
echo ============================================================
echo   مكتب الأستاذ بوخاري مريم - أداة إشهار عقاري V2 (المكتبي)
echo ============================================================

set "EXE_PATH=dist_executables\v2_modern\Ijraa_Ichhar_Notary_v2.exe"

if exist "%EXE_PATH%" (
    echo [OK] تشغيل البرنامج التنفيذي المباشر...
    start "" "%EXE_PATH%"
) else (
    echo [تنبيه] تشغيل البرنامج عبر بيئة بايثون مباشرة...
    python desktop_app\main_gui.py
)
pause

