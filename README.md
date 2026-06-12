# שרת עסקי - דשבורד + טופס אישור קבלה

## מה יש בקבצים
- `server.js` - השרת הראשי
- `package.json` - רשימת התלויות
- `public/dashboard.html` - הדשבורד (נתוני הזמנות/שירות)
- `public/confirmation-form.html` - טופס אישור קבלה לנהג

## שלב 1 - העלאה ל-GitHub
1. צור repository חדש ב-github.com (פרטי או ציבורי)
2. העלה את כל הקבצים והתיקיה `public` כמו שהם

## שלב 2 - חיבור ל-Render
1. כנס ל-render.com → New → Web Service
2. חבר את ה-repository שיצרת
3. Build Command: `npm install`
4. Start Command: `npm start`

## שלב 3 - משתני סביבה (Environment Variables) ב-Render
| NAME | VALUE |
|---|---|
| `GOOGLE_SERVICE_ACCOUNT_JSON` | אותו JSON שכבר השתמשנו ב-paypal-webhook |
| `SHEET_ID` | מזהה הגיליון (1_nM8UUSR8VqnhAoFUx-T2aiw5NnNt8tMQATRIIQ6Ib4) |
| `CONFIRMATIONS_FOLDER_ID` | מזהה תיקיית דרייב לאישורי קבלה (צריך ליצור תיקייה ולהעתיק את ה-ID מה-URL) |
| `CONFIRMATIONS_SHEET_NAME` | (אופציונלי) שם הלשונית, ברירת מחדל "אישורי קבלה" |

## שלב 4 - שיתוף הגיליון עם ה-Service Account
חשוב: יש לוודא שה-Service Account (האימייל מהקובץ JSON, מסתיים ב-`@insulation-494017.iam.gserviceaccount.com`) הוא **Editor** על:
1. הגיליון (`SHEET_ID`)
2. תיקיית הדרייב (`CONFIRMATIONS_FOLDER_ID`)

(Share → הדבק את כתובת המייל של ה-Service Account → Editor)

## כתובות לאחר ההעלאה
- דשבורד: `https://<your-app>.onrender.com/dashboard`
- טופס אישור קבלה: `https://<your-app>.onrender.com/confirmation-form`
