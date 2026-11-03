@echo off
chcp 65001 > nul
echo ============================================================
echo   مكتب الأستاذ بوخاري مريم - أداة إشهار عقاري V1 (الكلاسيكي)
echo ============================================================

set "EXE_PATH=dist_executables\v1_classic\Ijraa_Ichhar_Notary.exe"

if exist "%EXE_PATH%" (
    echo [OK] تشغيل البرنامج الكلاسيكي V1...
    start "" "%EXE_PATH%"
) else (
    echo [خطأ] لم يتم العثور على %EXE_PATH%
)
pause

