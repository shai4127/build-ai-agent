# build-ai-agent

סוכני AI מבוססי תיקייה.

## flight-hunter: צייד הטיסות ✈️
מעקב אחרי המחיר הזול ביותר לטיסה ישירה תל אביב ⇄ הרקליון (15–21.8.2027), שלוש פעמים ביום. התוצאות נרשמות לטבלת Excel, והתראת WhatsApp נשלחת כשהמחיר יורד.

### התקנה על שולחן העבודה (Windows, PowerShell)
```powershell
$Desktop = [Environment]::GetFolderPath('Desktop')
git clone --branch claude/wizardly-brown-c7i5qf https://github.com/shai4127/build-ai-agent.git "$env:TEMP\build-ai-agent"
Copy-Item -Recurse "$env:TEMP\build-ai-agent\flight-hunter" "$Desktop\flight-hunter"
```
אחר כך פותחים את `flight-hunter` שעל שולחן העבודה ב-Claude Code וכותבים: **אני רוצה להתחיל**
