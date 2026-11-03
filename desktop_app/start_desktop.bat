@echo off
chcp 65001 > nul
echo تشغيل تطبيق إشهار عقاري المكتبي (V2)...
if exist "dist\Ijraa_Ichhar_Notary_v2\Ijraa_Ichhar_Notary_v2.exe" (
    start "" "dist\Ijraa_Ichhar_Notary_v2\Ijraa_Ichhar_Notary_v2.exe"
) else if exist "dist_executable\Ijraa_Ichhar_Notary_v2.exe" (
    start "" "dist_executable\Ijraa_Ichhar_Notary_v2.exe"
) else (
    echo جاري تشغيل النسخة عبر بايثون مباشرة...
    python main_gui.py
)
