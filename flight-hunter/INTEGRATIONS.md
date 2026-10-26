# חיבורים ומצב הפעלה

הסטטוסים האפשריים: `ACTIVE` (פעיל ומאומת), `READY_TO_ACTIVATE` (מוכן ומחכה לאישור), `SETUP_REQUIRED` (דורש פעולה), `UNAVAILABLE` (לא נתמך).
הסוכן מעדכן את הטבלה הזאת רק אחרי אימות בפועל. קיום של קובץ לא נחשב אימות.

| רכיב | סטטוס | תפקיד | אימות אחרון | פעולה נדרשת |
|---|---|---|---|---|
| Python 3.10+ ו-pip | SETUP_REQUIRED | הרצת הסקריפטים | טרם אומת על המחשב | להתקין מ-python.org ולסמן "Add to PATH" |
| ספריות (openpyxl, playwright, tzdata) ו-Chromium | SETUP_REQUIRED | דפדפן אוטומטי וכתיבה ל-Excel | נבדק בסביבת הבנייה (Linux) | `tools\setup.ps1` |
| Claude Code CLI (`claude`) | SETUP_REQUIRED | הרצה לא מפוקחת (`claude -p`) | טרם אומת | לוודא ש-`claude --version` עובד ב-PowerShell |
| WhatsApp דרך CallMeBot | SETUP_REQUIRED | התראה בשפל מחיר | טרם הוגדר | ראה הוראות למטה |
| Windows Task Scheduler: 3 משימות | READY_TO_ACTIVATE | בדיקות 00:00, 12:00 ו-20:00 UTC | לא רשומות | רק אחרי אישור בשיחת "אני רוצה להתחיל" |
| Google Flights / Skyscanner / Kayak | OPTIONAL (לפי זמינות) | מקורות מחיר | ראה RUN_LOG.md | אין. אתר חסום נרשם בהערות |

## הגדרת CallMeBot (פעם אחת, כ-2 דקות)
1. שמור באנשי הקשר בטלפון את המספר של CallMeBot. **המספר העדכני מופיע ב-https://www.callmebot.com/blog/free-api-whatsapp-messages/**, כי הוא משתנה מדי פעם.
2. שלח לאיש הקשר הזה ב-WhatsApp את ההודעה: `I allow callmebot to send me messages`
3. תקבל בחזרה הודעה עם **apikey**.
4. במחשב, בתיקיית הסוכן, הרץ `notepad .env` ומלא:
   - `CALLMEBOT_PHONE=972...` (המספר שלך בפורמט בינלאומי, בלי + ובלי 0 בהתחלה)
   - `CALLMEBOT_APIKEY=...`
5. בדוק עם `python tools/notify_whatsapp.py "בדיקה מצייד הטיסות"`. ההודעה אמורה להגיע לטלפון תוך דקה.

**פרטיות:** הקובץ `.env` לא נכנס ל-git, והסוכן חסום מלקרוא אותו (`.claude/settings.json`). CallMeBot יכול לשלוח רק למספר שרשום אצלו, כלומר רק אליך.

## פעולות שתמיד דורשות אישור נפרד
- רישום, שינוי או מחיקה של משימות מתוזמנות
- שינוי פרטי הטיסה או המקורות
- כל פעולה שאינה קריאה באתר חיצוני. **הזמנה ותשלום לעולם לא מתבצעים.**
