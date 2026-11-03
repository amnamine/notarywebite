@echo off
chcp 65001 > nul
echo ============================================================
echo   مكتب الأستاذ بوخاري مريم - أداة إشهار عقاري
echo ============================================================

if not exist "venv" (
    echo [1/3] إنشاء البيئة الافتراضية venv...
    python -m venv venv
    call venv\Scripts\activate.bat
    echo [2/3] تثبيت المتطلبات...
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate.bat
)

echo [3/3] تشغيل الخادم...
start "" http://127.0.0.1:5000
python app.py
pause

