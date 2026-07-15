# 🚀 Gemini-Outlook-AppBot


An automated, AI-powered recruitment assistant written in Python. The bot reads target company information from an Excel database, leverages **Gemini 2.5 Flash** (via the official `google-genai` SDK) to generate highly personalized, tailored cover letters, and automates local **MS Outlook** (via COM interface) to create ready-to-review drafts with your PDF resume attached automatically.

### ✨ Features
* **AI-Generated Cover Letters:** Dynamically tailors professional cover letters based on the company name and specific notes/requirements.
* **Outlook Desktop Automation:** Utilizes Windows COM integration (`pywin32`) to create, format, and save drafts directly inside your local MS Outlook client.
* **Smart Account Binding:** Automatically detects and binds to your preferred sending email account configured in Outlook.
* **Rate-Limit & Spam Prevention:** Incorporates built-in delays (6 seconds between applications) to respect Gemini API free tier limits and mail server policies.
* **Secure Environment:** Sensitive credentials (like API keys and emails) are strictly managed using a `.env` file and excluded from version control.

### 🛠️ Tech Stack & Dependencies
* **Language:** Python 3.x
* **AI Integration:** Google GenAI SDK (`google-genai`)
* **Automation:** `pywin32` (Windows COM Active X)
* **Data Handling:** `pandas`, `openpyxl` (for Excel manipulation)
* **Configuration:** `python-dotenv`

---


مساعد ذكي ومطور لأتمتة التقديم على الوظائف باستخدام بايثون والذكاء الاصطناعي. يقوم الروبوت بقراءة بيانات الشركات المستهدفة من ملف إكسل، ثم يتصل بنموذج **Gemini 2.5 Flash** لتوليد رسائل تغطية (Cover Letters) مخصصة واحترافية بناءً على متطلبات كل وظيفة، وفي النهاية يقوم بالتحكم في برنامج **MS Outlook** محلياً لإنشاء مسودات (Drafts) جاهزة للمراجعة مع إرفاق السيرة الذاتية PDF تلقائياً.

### ✨ المميزات
* **التخصيص الفائق بالذكاء الاصطناعي:** كتابة رسائل تغطية فريدة واحترافية لكل شركة بناءً على اسمها والملاحظات المحددة للوظيفة.
* **أتمتة كاملة لـ Outlook:** استخدام مكتبة `pywin32` للتحكم في تطبيق الـ Outlook المكتبي على ويندوز لإنشاء وحفظ المسودات مباشرة دون تدخل بشري.
* **الربط الذكي بالحساب المرسل:** إمكانية تحديد البريد الإلكتروني المرسل بدقة ليتطابق مع حساباتك النشطة في Outlook.
* **تجنب الحظر والـ Rate Limits:** وضع مهلة زمنية ذكية (6 ثوانٍ) بين كل طلب لضمان استقرار الاتصال بـ API جوجل وتجنب فلاتر السبام.
* **بيئة عمل آمنة:** فصل كامل للبيانات الحساسة مثل مفاتيح الـ API والإيميلات في ملف خارجي غير مكشوف `.env`.

### 🛠️ التقنيات والمكتبات المستخدمة
* **لغة البرمجة:** Python 3.x
* **الذكاء الاصطناعي:** حزمة جوجل الرسمية `google-genai`
* **الأتمتة وتكامل ويندوز:** `pywin32`
* **معالجة البيانات:** `pandas` و `openpyxl`
* **إدارة الإعدادات:** `python-dotenv`

---

## 🚀 How to Run / طريقة التشغيل

### 1. Prerequisites / المتطلبات الأساسية
Make sure you have **Windows OS** and **MS Outlook Desktop App** installed, then run:
تأكد من العمل على نظام ويندوز ووجود برنامج Outlook المكتبي، ثم قم بتثبيت المكتبات التالية:

```bash
pip install -r requirements.txt
