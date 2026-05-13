# דרישות בדיקות ואיכות — צוות 1
**חברי הצוות:** נדב אילוז, נדב מגן  
**חלק בפיתוח:** מסך התחברות, רישום עובדים, ניהול לקוחות  
**שפת פיתוח:** ‎C#‎ (Windows Forms) • שמירת נתונים: Excel דרך ClosedXML

---

## 1. שני CFG (Control Flow Graphs)

### 1.1 פונקציה ראשונה: `Validators.IsValidUsername`

```csharp
public static bool IsValidUsername(string username, out string error)
{
    error = "";                                              // N1
    if (string.IsNullOrWhiteSpace(username))                 // N2
    {
        error = "Username is required.";
        return false;                                        // N3 -> EXIT
    }
    if (username.Length < 6 || username.Length > 8)          // N4
    {
        error = "Username must be 6-8 characters.";
        return false;                                        // N5 -> EXIT
    }
    int digitCount = 0;                                      // N6
    foreach (char c in username)                             // N7 (loop)
    {
        if (char.IsDigit(c))                                 // N8
            digitCount++;                                    // N9
        else if (!IsEnglishLetter(c))                        // N10
        {
            error = "Username may contain only English letters and digits.";
            return false;                                    // N11 -> EXIT
        }
    }
    if (digitCount > 2)                                      // N12
    {
        error = "Username may contain at most 2 digits.";
        return false;                                        // N13 -> EXIT
    }
    return true;                                             // N14 -> EXIT
}
```

**צמתים (Nodes):**

| צומת | תיאור |
|---|---|
| N1 | אתחול `error = ""` |
| N2 | תנאי: `string.IsNullOrWhiteSpace(username)` |
| N3 | החזרת false (שם משתמש ריק) |
| N4 | תנאי: אורך שם המשתמש מחוץ לטווח 6–8 |
| N5 | החזרת false (אורך לא תקין) |
| N6 | אתחול `digitCount = 0` |
| N7 | ראש לולאה: `foreach c in username` |
| N8 | תנאי: `char.IsDigit(c)` |
| N9 | `digitCount++` |
| N10 | תנאי: `!IsEnglishLetter(c)` |
| N11 | החזרת false (תו לא חוקי) |
| N12 | תנאי: `digitCount > 2` |
| N13 | החזרת false (יותר מ-2 ספרות) |
| N14 | החזרת true (תקין) |
| EXIT | סיום הפונקציה |

**קשתות (Edges):**

```
N1 → N2
N2 → N3   (true)
N2 → N4   (false)
N3 → EXIT
N4 → N5   (true)
N4 → N6   (false)
N5 → EXIT
N6 → N7
N7 → N8   (יש עוד תו)
N7 → N12  (סיום הלולאה)
N8 → N9   (true)
N8 → N10  (false)
N9 → N7
N10 → N11 (true)
N10 → N7  (false)
N11 → EXIT
N12 → N13 (true)
N12 → N14 (false)
N13 → EXIT
N14 → EXIT
```

> **הערה לציור:** ב-Word / PowerPoint / draw.io ציירו צמתים עגולים עם המספרים, וקשתות מסומנות true/false היכן שצריך. Cyclomatic Complexity = E − N + 2 = 19 − 14 + 2 = **7**.

---

### 1.2 פונקציה שנייה: `Validators.IsValidPassword`

```csharp
public static bool IsValidPassword(string password, out string error)
{
    error = "";                                              // M1
    if (string.IsNullOrEmpty(password))                      // M2
    {
        error = "Password is required.";
        return false;                                        // M3 -> EXIT
    }
    if (password.Length < 8 || password.Length > 10)         // M4
    {
        error = "Password must be 8-10 characters.";
        return false;                                        // M5 -> EXIT
    }
    bool hasLetter = false;                                  // M6
    bool hasDigit = false;
    bool hasSpecial = false;
    foreach (char c in password)                             // M7 (loop)
    {
        if (char.IsLetter(c)) hasLetter = true;              // M8 -> M9
        else if (char.IsDigit(c)) hasDigit = true;           // M10 -> M11
        else if (c == '!' || c == '#' || c == '$')           // M12
            hasSpecial = true;                               // M13
    }
    if (!hasLetter || !hasDigit || !hasSpecial)              // M14
    {
        error = "Password must contain ...";
        return false;                                        // M15 -> EXIT
    }
    return true;                                             // M16 -> EXIT
}
```

**צמתים:**

| צומת | תיאור |
|---|---|
| M1 | אתחול error |
| M2 | תנאי: סיסמה ריקה |
| M3 | החזרת false |
| M4 | תנאי: אורך סיסמה מחוץ ל-8–10 |
| M5 | החזרת false |
| M6 | אתחול 3 דגלים: hasLetter/hasDigit/hasSpecial |
| M7 | ראש לולאה: foreach c |
| M8 | תנאי: `char.IsLetter(c)` |
| M9 | `hasLetter = true` |
| M10 | תנאי: `char.IsDigit(c)` |
| M11 | `hasDigit = true` |
| M12 | תנאי: c ∈ {!, #, $} |
| M13 | `hasSpecial = true` |
| M14 | תנאי: לא קיים אחד מ-3 הדגלים |
| M15 | החזרת false |
| M16 | החזרת true |

**קשתות:**

```
M1→M2 ; M2→M3(t) ; M2→M4(f) ; M3→EXIT
M4→M5(t) ; M4→M6(f) ; M5→EXIT
M6→M7
M7→M8 (יש תו)
M7→M14 (סיום)
M8→M9(t) ; M8→M10(f)
M9→M7
M10→M11(t) ; M10→M12(f)
M11→M7
M12→M13(t) ; M12→M7(f)
M13→M7
M14→M15(t) ; M14→M16(f)
M15→EXIT ; M16→EXIT
```

Cyclomatic Complexity = E − N + 2 = 22 − 16 + 2 = **8**.

---

## 2. שלושה סיפורי משתמש (User Stories)

### US-1 — רישום עובד חדש
> **בתור** מזכירה חדשה במרפאה,  
> **אני רוצה** להירשם למערכת בעזרת שם משתמש, סיסמה, מספר עובד, מייל, ת"ז ותפקיד,  
> **כדי שאוכל** להתחבר בהמשך ולנהל לקוחות.

קריטריונים לקבלה:
- שם משתמש 6–8 תווים, עד 2 ספרות, השאר אותיות באנגלית.
- סיסמה 8–10 תווים, עם לפחות אות, ספרה ותו מיוחד (`!`, `#`, `$`).
- מספר עובד = 4 ספרות בדיוק.
- ת"ז 9 ספרות עם ספרת ביקורת תקינה (אלגוריתם ת"ז ישראלית).
- בלחיצה על "Register" העובד נשמר בקובץ ה-Excel.

### US-2 — התחברות
> **בתור** וטרינר רשום,  
> **אני רוצה** להתחבר עם שם משתמש וסיסמה,  
> **כדי לקבל** גישה לתפריט הראשי בהתאם לתפקידי.

קריטריונים:
- צירוף תקין → פתיחת `MainForm` עם שם המשתמש בתפריט.
- צירוף לא תקין → הודעת שגיאה והשארה במסך ההתחברות.

### US-3 — חיפוש לקוח לפי ת"ז או טלפון
> **בתור** מזכירה,  
> **אני רוצה** להזין ת"ז או טלפון ולבחור את אופן החיפוש,  
> **כדי למצוא** במהירות לקוח קיים ולראות את החיות שבבעלותו.

קריטריונים:
- בחירה "National ID" → סינון לפי הכלת המחרוזת בשדה הת"ז.
- בחירה "Phone" → סינון לפי הכלת המחרוזת בשדה הטלפון.
- לחיצה על שורה מציגה ברשימה הימנית את החיות שבבעלות הלקוח.

---

## 3. ארבעה מקרי בדיקה (משני סיפורי משתמש שונים)

### מקרי בדיקה הנגזרים מ-US-1 (רישום עובד)

| # | TC ID | תיאור | קלט | תוצאה צפויה | תוצאה בפועל |
|---|---|---|---|---|---|
| 1 | TC-1.1 | רישום תקין של מזכירה | user=`reception`, pass=`Clinic$99`, emp=`1001`, name=`Dana Levi`, mail=`d@x.co.il`, id=`123456782`, role=Secretary | חלון "Registered" מופיע, העובדת מופיעה ב-`Employees` ב-Excel | כצפוי |
| 2 | TC-1.2 | סיסמה ללא תו מיוחד | user=`vetuser`, pass=`Pass1234` (אורך 8, חסר תו מיוחד) | הודעת שגיאה: "Password must contain at least one letter, one digit and one special character" | כצפוי |

### מקרי בדיקה הנגזרים מ-US-3 (חיפוש לקוח)

| # | TC ID | תיאור | קלט | תוצאה צפויה | תוצאה בפועל |
|---|---|---|---|---|---|
| 3 | TC-3.1 | חיפוש לפי ת"ז קיימת | קיים לקוח עם ת"ז `123456782`. מקלידים `123456782` ובוחרים "National ID" | הטבלה מציגה שורה אחת – הלקוח | כצפוי |
| 4 | TC-3.2 | חיפוש לפי טלפון שאינו קיים | קיים לקוח עם טלפון `0501234567`. מקלידים `0509999999`, "Phone" | הטבלה ריקה, אין שגיאה | כצפוי |

---

## 4. שתי בדיקות פונקציונאליות + שתי בדיקות GUI

### 4.1 בדיקות פונקציונאליות (Functional)

**FT-1: ולידציית סיסמה**  
מה נבדק: לוגיקת `IsValidPassword` (כללי 8–10 תווים, אות+ספרה+תו מיוחד).  
איך: הזנו 5 סיסמאות שונות בטופס Register (קצרה, ארוכה מדי, ללא ספרה, ללא תו מיוחד, תקינה).  
תוצאה: 4 הראשונות נדחו עם הודעת שגיאה ספציפית, החמישית עברה בהצלחה. ✅

**FT-2: שיוך חיות ללקוח**  
מה נבדק: `CustomersForm` מציגה ברשימה הימנית את כל החיות עם `OwnerNationalId` השווה ל-`NationalId` של הלקוח שנבחר.  
איך: יצרנו לקוח, הוספנו 2 חיות שלו ועוד 1 של לקוח אחר, ובחרנו את הלקוח בטבלה.  
תוצאה: הופיעו רק 2 החיות שלו. ✅

### 4.2 בדיקות GUI

**GT-1: השדה Password מוסתר**  
מה נבדק: ב-`LoginForm` וב-`RegisterForm` תווי הסיסמה מוצגים כנקודות (`UseSystemPasswordChar = true`).  
איך: פתחנו כל אחד מהמסכים, הקלדנו `abc123!`.  
תוצאה: הוצגו `•••••••`. ✅

**GT-2: כפתור "← Main Menu" סוגר את `CustomersForm` ומחזיר ל-`MainForm`**  
מה נבדק: הפונקציונליות של ניווט חזרה.  
איך: התחברנו, פתחנו את מסך הלקוחות, לחצנו על הכפתור.  
תוצאה: החלון נסגר, `MainForm` מקבל מיקוד. ✅

---

## 5. תרחיש בדיקה + 2 מקרי בדיקה + תסריטי בדיקה

### תרחיש בדיקה (Test Scenario)
**TS-LOGIN: אימות זהות בכניסה למערכת.**  
המשתמש מזין שם משתמש וסיסמה ב-`LoginForm`. המערכת בודקת מול רשימת `Employees` שב-Excel: אם יש התאמה — נפתח `MainForm`; אחרת מוצגת הודעת שגיאה.

### מקרי בדיקה הנגזרים מהתרחיש

| TC ID | תיאור | תוצאה צפויה |
|---|---|---|
| TC-LOGIN-1 | התחברות עם פרטים נכונים של עובד שנרשם | פתיחת `MainForm` עם שם המשתמש בכותרת |
| TC-LOGIN-2 | התחברות עם סיסמה שגויה | הודעה "Invalid username or password", שדה הסיסמה מתאפס ומקבל focus |

### תסריטי בדיקה (Test Scripts)

#### תסריט TC-LOGIN-1

| שלב | פעולה | קלט | תוצאה צפויה |
|---|---|---|---|
| 1 | הפעלת `ClinicVets.exe` | – | מסך התחברות נפתח |
| 2 | הקלדה בשדה Username | `vetadmin` | האותיות מופיעות בשדה |
| 3 | הקלדה בשדה Password | `Clinic$99` | מוצגות 9 נקודות |
| 4 | לחיצה על Login | – | המסך נסגר, `MainForm` נפתח עם הכותרת "Welcome, Dana Levi" |

#### תסריט TC-LOGIN-2

| שלב | פעולה | קלט | תוצאה צפויה |
|---|---|---|---|
| 1 | הפעלת `ClinicVets.exe` | – | מסך התחברות נפתח |
| 2 | הקלדה בשדה Username | `vetadmin` | האותיות מופיעות |
| 3 | הקלדה בשדה Password | `WrongPass1!` | מוצגות 11 נקודות |
| 4 | לחיצה על Login | – | חלון שגיאה "Invalid username or password" |
| 5 | לחיצה על OK בחלון השגיאה | – | חלון השגיאה נסגר, שדה Password ריק וב-focus |

---

## איך להגיש למורה

1. פותחים את הקובץ הזה ב-Word.
2. הופכים את ה-Markdown לעיצוב Word (כותרות, טבלאות) — אפשר להעתיק ולהדביק.
3. ציירו את שני ה-CFG ב-draw.io / PowerPoint על פי רשימות הצמתים והקשתות שלמעלה, ושמרו כתמונה בתוך המסמך.
4. צרפו צילומי מסך של ההרצות (Register תקין, שגיאת סיסמה, התחברות, חיפוש לקוח, הצגת חיות הלקוח, ניווט בכפתור Main Menu).
5. שמרו כ-PDF והגישו ביחד עם קובץ ה-`.exe` וקישור לתיקייה המשותפת.
