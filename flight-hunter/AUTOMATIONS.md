# משימות מתוזמנות

> ⚠️ ההגדרות בקובץ הזה **אינן פעילות** עד שהן נרשמות ב-Windows Task Scheduler (`tools/register_tasks.ps1`) ומאומתות עם `Get-ScheduledTask`. המצב בפועל מתועד ב-[INTEGRATIONS.md](INTEGRATIONS.md).

| משימה | מועד (UTC) | שעון ישראל, חורף / קיץ | שם ב-Windows |
|---|---|---|---|
| בדיקה 1 | כל יום 00:00 | 02:00 / 03:00 | `\FlightHunter\FlightHunter-0000UTC` |
| בדיקה 2 | כל יום 12:00 | 14:00 / 15:00 | `\FlightHunter\FlightHunter-1200UTC` |
| בדיקה 3 | כל יום 20:00 | 22:00 / 23:00 | `\FlightHunter\FlightHunter-2000UTC` |

- **אזור זמן:** UTC, לבקשת הבעלים. הטריגרים נרשמים עם סיומת `Z`, ולכן לא זזים בין שעון חורף לקיץ.
- **תקופה:** מרגע ההפעלה ועד 15.8.2027, כולל. אחר כך הטריגרים פגים, והסקריפט עוצר בעצמו.
- **מה רץ:** Task Scheduler מפעיל את `powershell -File tools\run_check.ps1`, שמריץ `claude -p` לפי ההוראות ב-`automations/check-prices.md`.
- **תוצר:** שורה ב-`flights.xlsx`, יומן ב-`logs\run_<UTC>.log`, שורה ב-`RUN_LOG.md`, ו-WhatsApp רק בשפל חדש.
- **חיבורים נדרשים:** Python וספריות (`tools\setup.ps1`), Claude Code CLI מחובר (`claude` ב-PATH), WhatsApp Web מחובר (`python tools/notify_whatsapp.py --login`) ומספר ב-`.env`.
- **הרשאות לריצה לא מפוקחת:** רק הכלים שמופיעים ב-`--allowedTools` בתוך `run_check.ps1`: קריאה, הסקריפטים ב-`tools/`, WebSearch/WebFetch ועריכת `RUN_LOG.md`. אין גישה לכלים אחרים.

## כשלים
| מצב | התנהגות |
|---|---|
| המחשב כבוי בזמן הבדיקה | `StartWhenAvailable`: הבדיקה רצה כשהמחשב נדלק |
| המחשב במצב שינה | `WakeToRun`: המחשב מתעורר לבדיקה |
| ריצה קודמת עדיין פעילה | הריצה החדשה מדלגת (קובץ נעילה) |
| הריצה נמשכת יותר מ-40 דקות | Windows עוצר אותה |
| `flights.xlsx` פתוח ב-Excel | השורה נכנסת לתור ונכתבת בריצה הבאה |
| אתר חסום | ממשיכים לשאר האתרים ומתעדים |
| שליחת WhatsApp נכשלה (למשל החיבור פג) | הכישלון נרשם ביומן, והשורה בטבלה נשמרת כרגיל. אם מופיע `SETUP_REQUIRED`, צריך `--login` מחדש |

## ניהול
- **הפעלה:** `powershell -ExecutionPolicy Bypass -File tools\register_tasks.ps1`
- **עצירה:** `powershell -ExecutionPolicy Bypass -File tools\unregister_tasks.ps1`
- **בדיקה ידנית מיידית:** לכתוב בשיחה "תבדוק עכשיו".
