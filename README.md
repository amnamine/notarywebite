# أدوات مكتب التوثيق — إشهار عقاري (الإصدار الثاني V2)

تطبيق رسمي لتحضير العقود التوثيقية بصيغة Word (`.docx`) باللغة العربية مع الإطار الرسمي للجمهورية الجزائرية.

---

## المكونات المتوفرة وطرق التشغيل:

### 1. موقع الويب الداخلي (Flask Web App):
- **التشغيل المباشر**: اضغط نقرتين على [`start.bat`](file:///d:/AMINE2/COURS%20FAC/travaux/NOTARYY/notaryv2/notary-main/start.bat) أو نفّذ:
  ```bash
  python app.py
  ```
- **العنوان**: `http://127.0.0.1:5000`
- **حساب المدير الافتراضي**:
  - اسم المستخدم: `admin`
  - كلمة المرور: `admin123`
- **المميزات**: معالجة في المتصفح بالكامل (Client-side JS)، سرية تامة دون رفع ملفات، إدارة مستخدمين، وسجل عمليات نشط.

### 2. التطبيق المكتبي المستقل (Tkinter Desktop App V2):
- **ملف التشغيل التنفيذي الجاهز (.exe)**:
  - المسار: [`dist/Ijraa_Ichhar_Notary_v2/Ijraa_Ichhar_Notary_v2.exe`](file:///d:/AMINE2/COURS%20FAC/travaux/NOTARYY/notaryv2/notary-main/dist/Ijraa_Ichhar_Notary_v2/Ijraa_Ichhar_Notary_v2.exe)
  - أو اضغط نقرتين على: [`desktop_app/start_desktop.bat`](file:///d:/AMINE2/COURS%20FAC/travaux/NOTARYY/notaryv2/notary-main/desktop_app/start_desktop.bat)
- **المميزات**: يعمل دون الحاجة لتثبيت بايثون أو وجود اتصال بالإنترنت، استخراج آلي لبيانات مسح الأراضي وتاريخ العقد والفهرس.

### 3. النسخة المستقلة أحادية الملف (Offline USB Single File):
- متوفرة في المجلد [`dist/index_offline_single_file.html`](file:///d:/AMINE2/COURS%20FAC/travaux/NOTARYY/notaryv2/notary-main/dist/index_offline_single_file.html) للعمل على أي متصفح مباشرة دون خادم.
- يمكن إعادة بنائها في أي وقت عبر تشغيل:
  ```bash
  python build_single_file.py
  ```

---

## هيكل المجلدات المنظم (Clean Project Structure)
```
notary-main/
├── app.py                      # نقطة تشغيل خادم الويب Flask
├── start.bat                   # ملف التشغيل السريع للموقع (ويندوز)
├── build_single_file.py        # مولد نسخة الويب المستقلة للاستخدام دون إنترنت
├── requirements.txt            # متطلبات بايثون الأساسية
│
├── core_engine/                # المحرك البرمجي المشترك لمعالجة المستندات والـ XML
│   ├── ijraa_converter.py      # محرك إدراج العقود وتنسيق الجداول والأرقام والفواصل
│   ├── detect_box.py           # مستخرج مراجع المسح والفهرس والبلدية
│   ├── xml_helpers.py          # أدوات فحص وتعديل تراكيب الـ Word OpenXML
│   └── models.py               # قاعدة البيانات (المستخدمين وسجل العمليات)
│
├── desktop_app/                # تطبيق سطح المكتب (Tkinter Desktop V2)
│   ├── main_gui.py             # الكود المصدري للواجهة
│   ├── start_desktop.bat       # ملف تشغيل النسخة المكتبية
│   └── docx_templates/         # القوالب الرسمية المعتمدة
│
├── dist/                       # مجلد الملفات الجاهزة للتوزيع
│   ├── Ijraa_Ichhar_Notary_v2/ # تطبيق الويندوز التنفيذي (.exe) V2
│   └── index_offline_single_file.html # نسخة الويب أحادية الملف
│
├── routes/                     # مسارات ونقاط نهاية خادم الويب
├── static/                     # ملفات الجافاسكريبت والقوالب والمظهر (CSS/JS)
├── templates/                  # قوالب صفحات الويب (Jinja2 HTML)
└── tests/                      # الاختبارات التلقائية المعتمدة
```
