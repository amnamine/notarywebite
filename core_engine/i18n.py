"""
Dictionary of translations for Arabic, French, and English.
"""

TRANSLATIONS = {
    'ar': {
        # Header & Navigation
        'office_name': 'مكتب الأستاذة حاجي آسيا',
        'office_title': 'موثقة — الجزائر',
        'users': 'المستخدمين',
        'logs': 'سجل العمليات',
        'logout': 'خروج',
        'theme_toggle': 'تغيير المظهر',
        
        # Tool A - Ijraa Ichhar
        'page_title': 'إجــراء إشهار عقاري — أدوات تحضير العقود',
        'intro_title': 'إجــراء إشهار عقاري',
        'intro_desc': 'ضع العقد المكتوب (Word) فيوضع داخل إطار «إجــراء إشهار عقاري». الكلمات والأرقام لا تتغير، يتغيّر الشكل فقط.',
        
        'step1_title': 'ملف العقد',
        'drop_title': 'اسحب ملف Word هنا أو اضغط للاختيار',
        'drop_subtitle': 'صيغة .docx أو .doc',
        
        'step2_title': 'معلومات العقد ومراجع مسح الأراضي',
        'step2_hint': 'تُملأ تلقائياً من نص الملف. تأكد منها قبل الإنشاء.',
        'f_date': 'تاريخ تحرير العقد',
        'f_date_ph': 'مثال: 2026/09/02',
        'f_fihris': 'الفهرس',
        'f_fihris_ph': 'مثال: 2026/619',
        'f_qism': 'قسم',
        'f_qism_ph': 'مثال: 27',
        'f_bal': 'بلدية',
        'f_bal_ph': 'مثال: وادي العنب',
        'f_maj': 'مجموعة ملكية رقم',
        'f_maj_ph': 'مثال: 155',
        'f_hissa': 'حصة رقم',
        'f_hissa_ph': 'فارغة إن لم توجد',
        'f_wil': 'الولاية (الحفظ العقاري والمحافظ)',
        'f_wil_ph': 'مثال: عنابة',
        
        'step3_title': 'الإيداع والرسم',
        'step3_hint': 'اختياري. ما تتركه فارغاً يبقى نقاطاً «......» في الملف لتُكتب باليد.',
        'f_ida_hajm': 'إيداع — حجم',
        'f_ida_raqm': 'إيداع — رقم',
        'f_fi': 'في',
        'f_mujallad': 'مجلد',
        'f_muj_raqm': 'رقم (المجلد)',
        'f_rasm': 'رسم',
        
        'step4_title': 'الإنشاء والتحميل',
        'btn_generate': 'إنشاء الملف',
        'btn_download': 'تحميل «إجــراء إشهار عقاري»',
        
        'what_it_does': 'ما الذي يفعله؟',
        'rule_1': 'الترويسة، إطار الإيداع والرسم، مراجع مسح الأراضي، والعقد في العمود العريض مع «إطار مخصص للمحافظ العقاري»، والإطار ينزل حتى أسفل كل صفحة، وأرقام الصفحات.',
        'rule_2': 'الجداول تبقى كما هي في ملفك. إذا كان الجدول أعرض من الإطار يصغر كله بنفس النسبة.',
        'rule_3': 'حذف الأسطر الفارغة، والخطوط «-----» حتى آخر السطر في نهاية الفقرات (لا العناوين)، والأرقام مثل 91/148 كما كُتبت.',
        'rule_4': 'يقبل ملفات Word الجديدة .docx والقديمة .doc.',
        'privacy_note': 'يبقى الملف في جهازك. التحويل يتم مباشرة داخل المتصفح ولا يُرسل الملف أو يُحفظ في الخادم أبداً.',
        
        # Login Page
        'login_title': 'تسجيل الدخول',
        'login_desc': 'هذه الأداة خاصة بمكتب التوثيق. أدخل بيانات الدخول للمتابعة.',
        'username': 'اسم المستخدم',
        'password': 'كلمة المرور',
        'login_btn': 'دخول',
        
        # Admin Users
        'admin_users_title': 'إدارة مستخدمي المكتب',
        'add_user_title': 'إضافة موظف جديد',
        'role': 'الدور',
        'role_staff': 'موظف (Staff)',
        'role_admin': 'مسؤول (Admin)',
        'create_acc_btn': 'إنشاء الحساب',
        'col_username': 'اسم المستخدم',
        'col_role': 'الدور',
        'col_status': 'الحالة',
        'col_created': 'تاريخ الإنشاء',
        'col_actions': 'إجراءات',
        'status_active': 'مفعل',
        'status_inactive': 'معطل',
        'action_disable': 'تعطيل',
        'action_enable': 'تفعيل',
        
        # Admin Logs
        'admin_logs_title': 'سجل عمليات تحضير العقود (Activity Log)',
        'admin_logs_hint': 'ملاحظة: السجل يحتفظ فقط بأسماء الملفات وعدد الصفحات/التحذيرات، ولا يخزن أي محتوى ملفات لضمان السرية التامة.',
        'log_user': 'المستخدم',
        'log_tool': 'الأداة',
        'log_file': 'اسم الملف الأصلي',
        'log_counts': 'الصفحات / الجداول',
        'log_alerts': 'التنبيهات',
        'log_time': 'التوقيت',
        'log_empty': 'لا توجد عمليات مسجلة بعد.'
    },
    'fr': {
        # Header & Navigation
        'office_name': 'Étude Notariale Me Hadji Assia',
        'office_title': 'Notaire — Algérie',
        'users': 'Utilisateurs',
        'logs': 'Journal d’activité',
        'logout': 'Déconnexion',
        'theme_toggle': 'Changer de thème',
        
        # Tool A - Ijraa Ichhar
        'page_title': 'Publicité Foncière — Préparation d’Actes Notariés',
        'intro_title': 'Procédure de Publicité Foncière',
        'intro_desc': 'Insérez l’acte rédigé (Word) dans le cadre officiel de la publicité foncière. Le texte et les chiffres restent inchangés, seule la mise en page s’adapte.',
        
        'step1_title': '1. Document de l’acte',
        'drop_title': 'Glissez le fichier Word ici ou cliquez pour choisir',
        'drop_subtitle': 'Formats .docx ou .doc acceptés',
        
        'step2_title': '2. Références de l’acte et du cadastre',
        'step2_hint': 'Remplies automatiquement depuis le document. Vérifiez avant validation.',
        'f_date': 'Date de rédaction de l’acte',
        'f_date_ph': 'Ex: 2026/09/02',
        'f_fihris': 'Numéro du Répertoire',
        'f_fihris_ph': 'Ex: 2026/619',
        'f_qism': 'Section',
        'f_qism_ph': 'Ex: 27',
        'f_bal': 'Commune',
        'f_bal_ph': 'Ex: Oued El Aneb',
        'f_maj': 'Ilot de propriété N°',
        'f_maj_ph': 'Ex: 155',
        'f_hissa': 'Lot N°',
        'f_hissa_ph': 'Laisser vide si aucun',
        'f_wil': 'Wilaya (Conservation Foncière)',
        'f_wil_ph': 'Ex: Annaba',
        
        'step3_title': '3. Dépôt et Droits',
        'step3_hint': 'Optionnel. Les champs vides resteront sous forme de pointillés «......».',
        'f_ida_hajm': 'Dépôt — Volume',
        'f_ida_raqm': 'Dépôt — Numéro',
        'f_fi': 'En date du',
        'f_mujallad': 'Registre',
        'f_muj_raqm': 'Numéro de registre',
        'f_rasm': 'Droits / Taxes',
        
        'step4_title': '4. Génération et Téléchargement',
        'btn_generate': 'Générer le document',
        'btn_download': 'Télécharger le document prêt',
        
        'what_it_does': 'Fonctionnalités et règles officielles :',
        'rule_1': 'Mise en page officielle : en-têtes, cadre de conservation foncière, références cadastrales et foliotage automatique.',
        'rule_2': 'Préservation des tableaux : mise à l’échelle automatique et alignement précis dans le cadre.',
        'rule_3': 'Nettoyage des paragraphes vides, tirets de fin de ligne et conservation des formats numériques (ex. 91/148).',
        'rule_4': 'Prise en charge des formats Word récents (.docx) et anciens (.doc).',
        'privacy_note': 'Confidentialité totale : Vos fichiers ne quittent jamais votre ordinateur. Le traitement s’effectue localement dans le navigateur.',
        
        # Login Page
        'login_title': 'Connexion',
        'login_desc': 'Outil réservé au personnel de l’étude. Veuillez vous identifier.',
        'username': 'Nom d’utilisateur',
        'password': 'Mot de passe',
        'login_btn': 'Se connecter',
        
        # Admin Users
        'admin_users_title': 'Gestion des utilisateurs',
        'add_user_title': 'Ajouter un collaborateur',
        'role': 'Rôle',
        'role_staff': 'Collaborateur (Staff)',
        'role_admin': 'Administrateur',
        'create_acc_btn': 'Créer le compte',
        'col_username': 'Utilisateur',
        'col_role': 'Rôle',
        'col_status': 'Statut',
        'col_created': 'Date de création',
        'col_actions': 'Actions',
        'status_active': 'Actif',
        'status_inactive': 'Désactivé',
        'action_disable': 'Désactiver',
        'action_enable': 'Activer',
        
        # Admin Logs
        'admin_logs_title': 'Journal des opérations (Activity Log)',
        'admin_logs_hint': 'Remarque : Le journal n’enregistre que le nom des fichiers et les statistiques de traitement. Aucun contenu n’est sauvegardé.',
        'log_user': 'Utilisateur',
        'log_tool': 'Outil',
        'log_file': 'Fichier original',
        'log_counts': 'Paragraphes / Tableaux',
        'log_alerts': 'Avertissements',
        'log_time': 'Date & Heure',
        'log_empty': 'Aucune opération enregistrée pour le moment.'
    },
    'en': {
        # Header & Navigation
        'office_name': 'Office of Maitre Hadji Assia',
        'office_title': 'Notary — Algeria',
        'users': 'Users',
        'logs': 'Activity Logs',
        'logout': 'Sign Out',
        'theme_toggle': 'Toggle Theme',
        
        # Tool A - Ijraa Ichhar
        'page_title': 'Real Estate Registration — Deed Preparation',
        'intro_title': 'Real Estate Registration Procedure',
        'intro_desc': 'Place your written Word deed inside the official notarial frame. Text and numbers stay intact; only formatting is applied.',
        
        'step1_title': '1. Deed Document',
        'drop_title': 'Drop your Word file here or click to browse',
        'drop_subtitle': '.docx or .doc format supported',
        
        'step2_title': '2. Deed Info & Cadastral Details',
        'step2_hint': 'Automatically parsed from document text. Please verify before creating.',
        'f_date': 'Deed Execution Date',
        'f_date_ph': 'e.g. 2026/09/02',
        'f_fihris': 'Registry Index Number',
        'f_fihris_ph': 'e.g. 2026/619',
        'f_qism': 'Section',
        'f_qism_ph': 'e.g. 27',
        'f_bal': 'Municipality (Commune)',
        'f_bal_ph': 'e.g. Oued El Aneb',
        'f_maj': 'Property Block No.',
        'f_maj_ph': 'e.g. 155',
        'f_hissa': 'Lot No.',
        'f_hissa_ph': 'Leave empty if none',
        'f_wil': 'Province (Wilaya / Land Registry)',
        'f_wil_ph': 'e.g. Annaba',
        
        'step3_title': '3. Deposit & Fees',
        'step3_hint': 'Optional. Blank inputs remain dotted lines «......» for manual entry.',
        'f_ida_hajm': 'Deposit — Volume',
        'f_ida_raqm': 'Deposit — Number',
        'f_fi': 'Dated',
        'f_mujallad': 'Binder/Volume',
        'f_muj_raqm': 'Binder Number',
        'f_rasm': 'Fees / Taxes',
        
        'step4_title': '4. Creation & Download',
        'btn_generate': 'Create Document',
        'btn_download': 'Download Resulting Deed',
        
        'what_it_does': 'What does this tool do?',
        'rule_1': 'Places standard header, registrar deposit frame, cadastral coordinates, and automatic official pagination.',
        'rule_2': 'Keeps tables within frame margins, automatically scaling down wide tables while preserving contents.',
        'rule_3': 'Cleans empty lines, inserts hyphen leaders on paragraph endings, and preserves numbers (e.g. 91/148).',
        'rule_4': 'Supports modern .docx and legacy .doc files.',
        'privacy_note': '100% Private: Deeds are processed locally inside your browser and are never uploaded or stored on any server.',
        
        # Login Page
        'login_title': 'Sign In',
        'login_desc': 'Internal office tool. Please sign in to proceed.',
        'username': 'Username',
        'password': 'Password',
        'login_btn': 'Sign In',
        
        # Admin Users
        'admin_users_title': 'User Management',
        'add_user_title': 'Add New User',
        'role': 'Role',
        'role_staff': 'Staff Member',
        'role_admin': 'Administrator',
        'create_acc_btn': 'Create Account',
        'col_username': 'Username',
        'col_role': 'Role',
        'col_status': 'Status',
        'col_created': 'Created At',
        'col_actions': 'Actions',
        'status_active': 'Active',
        'status_inactive': 'Disabled',
        'action_disable': 'Disable',
        'action_enable': 'Enable',
        
        # Admin Logs
        'admin_logs_title': 'Deed Processing Activity Log',
        'admin_logs_hint': 'Notice: Only file names and processing metrics are recorded. Document contents are never saved or stored.',
        'log_user': 'User',
        'log_tool': 'Tool',
        'log_file': 'Original Filename',
        'log_counts': 'Paragraphs / Tables',
        'log_alerts': 'Warnings',
        'log_time': 'Timestamp',
        'log_empty': 'No activity logged yet.'
    }
}

def get_text(lang: str, key: str) -> str:
    lang_dict = TRANSLATIONS.get(lang) or TRANSLATIONS['ar']
    return lang_dict.get(key) or TRANSLATIONS['ar'].get(key, key)

