@echo off
chcp 65001 > nul
echo ============================================================
echo   مكتب الأستاذة حاجي آسيا - موثقة - إشهار عقاري V1 (الكلاسيكي)
echo ============================================================

set "EXE_PATH=dist_executables\v1_classic\Ijraa_Ichhar_Notary_v1.exe"

if exist "%EXE_PATH%" (
    echo [OK] تشغيل البرنامج المستقل V1 بنجاح...
    start "" "%EXE_PATH%"
) else (
    echo [خطأ] لم يتم العثور على %EXE_PATH%
)
