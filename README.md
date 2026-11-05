# build-ai-agent

סוכני AI מבוססי תיקייה.

## flight-hunter: צייד הטיסות ✈️
מעקב אחרי המחיר הזול ביותר לטיסה ישירה תל אביב ⇄ הרקליון (15–21.8.2027), כל שעה. התוצאות נרשמות לטבלת Excel. התראת WhatsApp נשלחת כשהמחיר יורד, ובכל בוקר נשלח סיכום של 3 המחירים הזולים.

### מיקום הסוכן במחשב
`C:\ai_skills\ai-agent\flight-hunter`

### התקנה חדשה (Windows, PowerShell)
```powershell
git clone --branch claude/wizardly-brown-c7i5qf https://github.com/shai4127/build-ai-agent.git "$env:TEMP\build-ai-agent"
Copy-Item -Recurse "$env:TEMP\build-ai-agent\flight-hunter" "C:\ai_skills\ai-agent\flight-hunter"
```
אחר כך פותחים את `C:\ai_skills\ai-agent\flight-hunter` ב-Claude Code וכותבים: **אני רוצה להתחיל**

### עדכון תיקייה קיימת
```powershell
$Temp = "$env:TEMP\fh-update"; if (Test-Path $Temp) { Remove-Item -Recurse -Force $Temp }
git clone --depth 1 --branch claude/wizardly-brown-c7i5qf https://github.com/shai4127/build-ai-agent.git $Temp
powershell -ExecutionPolicy Bypass -File "$Temp\flight-hunter\tools\update.ps1" -Target "C:\ai_skills\ai-agent\flight-hunter"
```
העדכון לא נוגע ב-`flights.xlsx`, `.env`, בחיבור ל-WhatsApp, ב-`logs`, ב-`RUN_LOG.md`, ב-`MEMORY.md` וב-`INTEGRATIONS.md`. אם הטבלה בפורמט ישן או נערכה ידנית, היא מותאמת לעמודות החדשות לפי שמות העמודות. עמודות שהוספת בעצמך נשמרות בסוף הטבלה, לשוניות אחרות לא משתנות, ועותק גיבוי נשמר לפני השינוי. אחר כך פותחים את התיקייה ב-Claude Code וכותבים: **אני רוצה להתחיל**
