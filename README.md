# build-ai-agent

סוכני AI מבוססי תיקייה.

## flight-hunter: צייד הטיסות ✈️
מעקב אחרי המחיר הזול ביותר לטיסה ישירה תל אביב ⇄ הרקליון (15–21.8.2027), כל שעה. התוצאות נרשמות לטבלת Excel. התראת WhatsApp נשלחת כשהמחיר יורד, ובכל בוקר נשלח סיכום של 3 המחירים הזולים.

### התקנה על שולחן העבודה (Windows, PowerShell)
```powershell
$Desktop = [Environment]::GetFolderPath('Desktop')
git clone --branch claude/wizardly-brown-c7i5qf https://github.com/shai4127/build-ai-agent.git "$env:TEMP\build-ai-agent"
Copy-Item -Recurse "$env:TEMP\build-ai-agent\flight-hunter" "$Desktop\flight-hunter"
```
אחר כך פותחים את `flight-hunter` שעל שולחן העבודה ב-Claude Code וכותבים: **אני רוצה להתחיל**

### עדכון תיקייה קיימת
אם התיקייה כבר על שולחן העבודה, מריצים ב-PowerShell:
```powershell
$Desktop = [Environment]::GetFolderPath('Desktop')
$Temp = "$env:TEMP\fh-update"; if (Test-Path $Temp) { Remove-Item -Recurse -Force $Temp }
git clone --depth 1 --branch claude/wizardly-brown-c7i5qf https://github.com/shai4127/build-ai-agent.git $Temp
powershell -ExecutionPolicy Bypass -File "$Temp\flight-hunter\tools\update.ps1" -Target "$Desktop\flight-hunter"
```
העדכון לא נוגע ב-`flights.xlsx`, `.env`, בחיבור ל-WhatsApp, ב-`logs`, ב-`RUN_LOG.md`, ב-`MEMORY.md` וב-`INTEGRATIONS.md`. אחר כך פותחים את התיקייה ב-Claude Code וכותבים: **אני רוצה להתחיל**
