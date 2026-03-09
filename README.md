Українською (UA)
Що це
Набір макросів PowerPoint (VBA) + Python-скрипт, який дозволяє за 1 дію: обрати учня, експортувати всі слайди презентації в PNG, зібрати PDF та відправити його в Telegram через бота.

Список учнів зберігається у students.json, а шлях до нього та останній вибраний профіль якості — у HKCU (через GetSetting/SaveSetting).

Можливості
Вибір отримувача та якості (Fast/HD/2K/4K) у формі frmStudents.
​

Експорт слайдів у PNG (slide_001.png, slide_002.png, …) у тимчасову папку %TEMP%.
​

Збір PDF з PNG через img2pdf або відправка готового PDF (режим “готовий шлях” у Python).
​

Відправка PDF в Telegram через Bot API sendDocument.
​

Toast-повідомлення Windows на успіх (показує “ім’я учня - профіль - файл.pdf”), а при помилці — MsgBox у PowerPoint.

CRUD-редактор учнів frmStudentsManage: додати/оновити/деактивувати/видалити, пошук, вибір шляху до students.json.

Як це працює (потік)
Кнопка Ribbon викликає ExportAndSendPDF (VBA).
​

frmStudents повертає studentId, “людську” мітку (ім’я/нотатка), і профіль якості.

VBA експортує всі слайди активної презентації у PNG у тимчасову папку.
​

VBA запускає Python (зазвичай pythonw.exe, без консолі), передає: input_path, student_id, --student-label, --profile, --log-path.

Python:

знаходить chat_id учня через students.json;
​

(якщо input_path — папка) збирає PDF з slide_*.png;
​

відправляє PDF в Telegram;
​

пише службові рядки в ppt_send_log.txt (наприклад PDF_SIZE_BYTES=...) для VBA;

показує toast на успіх.
​

Структура проекту
PowerPoint VBA
modMain — orchestration: UI → export → run Python → parse log → повідомлення.
​

frmStudents — вибір учня та профілю якості, відновлює останній профіль з HKCU.

frmStudentsManage — CRUD редактор students.json + вибір шляху до файлу.

modStudentsStore — завантаження/збереження students.json (UTF‑8), генерація u001/u002..., пошук/утиліти.
​

modSettings — AppTitle, шлях до students.json, останній профіль якості (HKCU).
​

modTextIO — читання/запис UTF‑8 через ADODB.Stream.
​

JsonConverter — VBA-JSON парсер.
​

vba-form-moderniser — модулі/класи для “modern” UI у формах (label-based buttons тощо).

Python
send_pdf.py — збір PDF, відправка в Telegram, логування, toast.
​

config.json — bot_token (токен Telegram-бота).
​

students.json — список учнів (id/name/chat_id/active/note).

logs/send_pdf.log — ротаційний лог Python.
​

Вимоги
Windows + Microsoft PowerPoint (VBA макроси).
​

Python 3.12 (або сумісний) встановлений локально.
​

Python-пакети:

requests (HTTP до Telegram);
​

img2pdf (PNG → PDF);
​

winotify (toast на успіх).
​

Приклад установки:

text
py -3.12 -m pip install requests img2pdf winotify
(Або через конкретний шлях до python.exe — як у твоїй інсталяції.)

Налаштування
config.json (поруч із send_pdf.py):

json
{ "bot_token": "123456:ABC..." }
send_pdf.py валідатором перевіряє, що токен є.
​

students.json (приклад запису):

json
[
  { "id": "u001", "name": "Нікіта Семерня", "chat_id": 123456789, "active": true, "note": "9 клас" }
]
Неактивні (active=false) не підтягуються в picker.

Шлях до students.json у VBA зберігається в HKCU і змінюється через frmStudentsManage.

Типові проблеми
Не показує учнів: перевір шлях до students.json у frmStudentsManage та валідність JSON.

Fail/exitCode ≠ 0: відкрий %TEMP%\ppt_send_log.txt і send_pdf.py\logs\send_pdf.log.

Не приходить в Telegram: перевір bot_token в config.json і chat_id у students.json (має бути числом).

English (EN)
What it is
A PowerPoint VBA + Python helper that lets you: pick a student, export all slides to PNG, build a PDF, and send it to Telegram via a bot — in one flow.

Students are stored in students.json; the path to that file and the last used quality profile are stored in HKCU via VBA GetSetting/SaveSetting.

Features
Recipient + quality picker (frmStudents) with presets: Fast/HD/2K/4K.
​

Exports slides as slide_001.png, slide_002.png, … to a %TEMP% folder.
​

Builds PDF from PNG via img2pdf, or sends an already built PDF path.
​

Sends PDF to Telegram using Bot API sendDocument.
​

Windows toast on success (student name - quality profile - file.pdf); MsgBox on failure.

Student manager (frmStudentsManage): add/update/deactivate/delete students, search/filter, change students.json path.
​

High-level flow
Ribbon button calls ExportAndSendPDF (VBA).
​

frmStudents returns studentId, studentLabel, and the selected quality profile.

VBA exports all slides of the active presentation into PNG images.
​

VBA starts Python (typically pythonw.exe, hidden) passing: input_path, student_id, --student-label, --profile, --log-path.

Python:

looks up chat_id in students.json;
​

builds a PDF from slide_*.png when needed;
​

posts the document to Telegram;
​

writes machine-readable output into ppt_send_log.txt for VBA parsing;

shows a toast notification on success.
​

Project layout
PowerPoint VBA
modMain — orchestration: UI → export → run Python → parse log → messaging.
​

frmStudents — recipient + quality picker, restores last profile from HKCU.

frmStudentsManage — CRUD for students.json + browse path UI.

modStudentsStore — load/save UTF‑8 JSON, generate ids, helpers.
​

modSettings — HKCU settings wrapper (paths, last profile, app title).
​

modTextIO — UTF‑8 read/write via ADODB.Stream.
​

JsonConverter — VBA JSON parser.
​

vba-form-moderniser — UI styling/modern button behavior for UserForms.

Python
send_pdf.py — build/send/log/toast.
​

config.json — Telegram bot token.
​

students.json — student roster.

logs/send_pdf.log — rotating Python log.
​

Requirements
Windows + Microsoft PowerPoint with macros enabled.
​

Python 3.12+ installed locally.
​

Python packages: requests, img2pdf, winotify.
​

Install:

text
py -3.12 -m pip install requests img2pdf winotify
Configuration
config.json next to send_pdf.py must contain bot_token.
​

students.json is a list of objects with id, name, chat_id, optional active, note. Inactive entries are not shown in the picker.

The students.json path is stored in HKCU and can be changed in frmStudentsManage.

Troubleshooting
No students shown: validate students.json path and JSON content.

Non-zero exit code: check %TEMP%\ppt_send_log.txt and logs/send_pdf.log.

Telegram send failed: verify bot_token and chat_id values.