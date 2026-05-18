using System.Globalization;
using ClinicVets.Models;
using ClosedXML.Excel;

namespace ClinicVets.Data;

/// <summary>
/// מחלקה זו אחראית על ניהול הנתונים של המרפאה הווטרינרית מול קובץ אקסל.
/// היא משמשת כמעין מסד נתונים פשוט (Data Store) שקורא וכותב את רשימות האובייקטים לגיליונות (Worksheets) שונים באקסל.
/// המחלקה משתמשת בספריית ClosedXML כדי לעבוד עם הקבצים.
/// </summary>
public class ExcelDataStore
{
    // משתנה פרטי השומר את הנתיב לקובץ האקסל שבו נשמרים הנתונים
    private readonly string _filePath;

    // רשימות (Properties) המחזיקות את כל הנתונים בזיכרון (In-Memory) בזמן ריצת התוכנית.
    // לכל אחת יש get פומבי (קריאה לכולם) ו-set פרטי (שינוי הרשימה עצמה אפשרי רק מתוך המחלקה).
    public List<Employee> Employees { get; private set; } = new();
    public List<Customer> Customers { get; private set; } = new();
    public List<Animal> Animals { get; private set; } = new();
    public List<Visit> Visits { get; private set; } = new();
    public List<Medication> Medications { get; private set; } = new();
    public List<string> Species { get; private set; } = new();

    // מערך קבוע שמכיל את סוגי החיות הדיפולטיביים (ברירת מחדל) למקרה שהמערכת עולה פעם ראשונה
    private static readonly string[] DefaultSpecies = { "Dog", "Cat", "Reptile", "Bird" };

    /// <summary>
    /// בנאי (Constructor) המחלקה. מקבל את נתיב הקובץ ושומר אותו במשתנה הפרטי.
    /// </summary>
    public ExcelDataStore(string filePath)
    {
        _filePath = filePath;
    }

    /// <summary>
    /// מתודה לטעינת הנתונים מקובץ האקסל אל תוך הרשימות בזיכרון.
    /// </summary>
    public void Load()
    {
        // אם קובץ האקסל לא קיים בנתיב המבוקש (למשל, הפעלה ראשונה של התוכנה)
        if (!File.Exists(_filePath))
        {
            SeedDefaults(); // הזנת נתוני ברירת מחדל
            Save();         // שמירת הקובץ החדש כדי שייווצר פיזית
            return;         // יציאה מהמתודה כי אין מה לקרוא מהקובץ
        }

        // אם הקובץ קיים, פותחים אותו לקריאה בעזרת ClosedXML. ה-using מבטיח שהקובץ ייסגר בסיום.
        using var wb = new XLWorkbook(_filePath);

        // קריאת הנתונים מכל גיליון (Worksheet) לתוך הרשימות המתאימות
        Employees = ReadEmployees(wb);
        Customers = ReadCustomers(wb);
        Animals = ReadAnimals(wb);
        Visits = ReadVisits(wb);
        Medications = ReadMedications(wb);
        Species = ReadSpecies(wb);

        // בדיקה האם חסרים נתוני חובה שצריך להשלים (למשל אם הקובץ היה ריק בחלקו)
        bool dirty = false; // דגל (Flag) שמסמן אם בוצעו שינויים שדורשים שמירה מחדש
        if (Medications.Count == 0) { SeedMedications(); dirty = true; } // אם אין תרופות, טען ברירת מחדל
        if (Species.Count == 0) { SeedSpecies(); dirty = true; }         // אם אין סוגי חיות, טען ברירת מחדל

        // אם הוספנו נתוני ברירת מחדל, נשמור את הקובץ המעודכן
        if (dirty) Save();
    }

    /// <summary>
    /// מתודה לשמירת כל הרשימות מהזיכרון אל תוך קובץ האקסל.
    /// דורסת את הקובץ הקיים עם הנתונים המעודכנים.
    /// </summary>
    public void Save()
    {
        using var wb = new XLWorkbook(); // יצירת חוברת עבודה (Workbook) חדשה בזיכרון

        // כתיבת כל אחת מהרשימות לגיליון (Worksheet) ייעודי
        WriteEmployees(wb);
        WriteCustomers(wb);
        WriteAnimals(wb);
        WriteVisits(wb);
        WriteMedications(wb);
        WriteSpecies(wb);

        wb.SaveAs(_filePath); // שמירת חוברת העבודה לקובץ הפיזי בנתיב המוגדר
    }

    /// <summary>
    /// מתודה המנסה להוסיף חיה חדשה למערכת.
    /// </summary>
    /// <param name="animal">אובייקט החיה להוספה</param>
    /// <param name="error">פרמטר out שמחזיר הודעת שגיאה במידה וההוספה נכשלה</param>
    /// <returns>True אם ההוספה הצליחה, False אם נכשלה</returns>
    public bool TryAddAnimal(Animal animal, out string error)
    {
        error = "";
        // בדיקה: האם קיימת כבר חיה במערכת עם אותו מספר שבב
        if (Animals.Any(a => a.ChipNumber == animal.ChipNumber))
        {
            error = "Chip number already exists.";
            return false; // הוספה נכשלה
        }

        Animals.Add(animal); // הוספת החיה לרשימה בזיכרון
        Save();              // שמירת השינוי לקובץ האקסל
        return true;         // הוספה הצליחה
    }

    /// <summary>
    /// מתודה המנסה להוסיף סוג חיה חדש לרשימת הסוגים (Species).
    /// </summary>
    public bool TryAddSpecies(string name, out string error)
    {
        error = "";
        var trimmed = name.Trim(); // הסרת רווחים מתחילת וסוף המחרוזת

        // וידוא שהשם אינו ריק
        if (string.IsNullOrEmpty(trimmed))
        {
            error = "Species name is required.";
            return false;
        }

        // וידוא שהסוג לא קיים כבר (תוך התעלמות מאותיות גדולות/קטנות)
        if (Species.Any(s => string.Equals(s, trimmed, StringComparison.OrdinalIgnoreCase)))
        {
            error = "Species already exists.";
            return false;
        }

        Species.Add(trimmed); // הוספה לרשימה
        Save();               // שמירה לאקסל
        return true;
    }

    /// <summary>
    /// מתודה למחיקת סוג חיה מהמערכת.
    /// </summary>
    public bool RemoveSpecies(string name)
    {
        // חיפוש האינדקס של הסוג ברשימה (תוך התעלמות מאותיות גדולות/קטנות)
        var idx = Species.FindIndex(s => string.Equals(s, name, StringComparison.OrdinalIgnoreCase));
        if (idx < 0) return false; // לא נמצא

        Species.RemoveAt(idx); // הסרה מהרשימה
        Save();                // שמירה לאקסל
        return true;
    }

    // --- מתודות Seed - הזנת נתוני התחלתיים (ברירת מחדל) למערכת ריקה ---

    private void SeedDefaults()
    {
        SeedMedications();
        SeedSpecies();
    }

    private void SeedSpecies()
    {
        // המרת מערך ה-DefaultSpecies לרשימה
        Species = DefaultSpecies.ToList();
    }

    private void SeedMedications()
    {
        // אתחול רשימת התרופות עם נתוני בסיס קבועים
        Medications = new List<Medication>
        {
            new() { Name = "Antibiotics", Price = 80m, StockQuantity = 50 },
            new() { Name = "Painkiller", Price = 45m, StockQuantity = 80 },
            new() { Name = "Vaccine - Annual", Price = 120m, StockQuantity = 30 },
            new() { Name = "Anti-parasitic", Price = 60m, StockQuantity = 40 },
        };
    }

    // ====================================================================
    // --- מתודות לקריאה וכתיבה של ישויות ספציפיות מאקסל ולאקסל (CRUD) ---
    // ====================================================================

    // ---------- Employees (עובדים) ----------

    private static List<Employee> ReadEmployees(XLWorkbook wb)
    {
        var list = new List<Employee>();
        // מנסה להשיג את הגיליון "Employees". אם לא קיים, מחזיר רשימה ריקה.
        if (!wb.TryGetWorksheet("Employees", out var ws)) return list;

        // עובר על כל השורות בגיליון שיש בהן נתונים, מדלג על השורה הראשונה (Skip(1)) כי היא שורת כותרות.
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            list.Add(new Employee
            {
                Username = row.Cell(1).GetString(),
                Password = row.Cell(2).GetString(),
                EmployeeNumber = row.Cell(3).GetString(),
                FullName = row.Cell(4).GetString(),
                Email = row.Cell(5).GetString(),
                NationalId = row.Cell(6).GetString(),
                // מנסה להמיר את המחרוזת לתצורת Enum של Role. אם נכשל, שם ברירת מחדל Secretary
                Role = Enum.TryParse<Role>(row.Cell(7).GetString(), out var r) ? r : Role.Secretary,
            });
        }
        return list;
    }

    private void WriteEmployees(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Employees"); // יצירת גיליון חדש
        // כתיבת שורת הכותרות בעזרת מתודת העזר
        WriteHeader(ws, "Username", "Password", "EmployeeNumber", "FullName", "Email", "NationalId", "Role");

        // מעבר על הרשימה וכתיבת הנתונים החל משורה 2 (כי שורה 1 היא כותרת)
        for (int i = 0; i < Employees.Count; i++)
        {
            var e = Employees[i];
            var row = ws.Row(i + 2);
            row.Cell(1).Value = e.Username;
            row.Cell(2).Value = e.Password;
            row.Cell(3).Value = e.EmployeeNumber;
            row.Cell(4).Value = e.FullName;
            row.Cell(5).Value = e.Email;
            row.Cell(6).Value = e.NationalId;
            row.Cell(7).Value = e.Role.ToString(); // שמירת ה-Enum כמחרוזת
        }
    }

    // ---------- Customers (לקוחות) ----------

    private static List<Customer> ReadCustomers(XLWorkbook wb)
    {
        var list = new List<Customer>();
        if (!wb.TryGetWorksheet("Customers", out var ws)) return list;
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            list.Add(new Customer
            {
                FullName = row.Cell(1).GetString(),
                NationalId = row.Cell(2).GetString(),
                Phone = row.Cell(3).GetString(),
                Email = row.Cell(4).GetString(),
            });
        }
        return list;
    }

    private void WriteCustomers(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Customers");
        WriteHeader(ws, "FullName", "NationalId", "Phone", "Email");
        for (int i = 0; i < Customers.Count; i++)
        {
            var c = Customers[i];
            var row = ws.Row(i + 2);
            row.Cell(1).Value = c.FullName;
            row.Cell(2).Value = c.NationalId;
            row.Cell(3).Value = c.Phone;
            row.Cell(4).Value = c.Email;
        }
    }

    // ---------- Animals (חיות) ----------

    private static List<Animal> ReadAnimals(XLWorkbook wb)
    {
        var list = new List<Animal>();
        if (!wb.TryGetWorksheet("Animals", out var ws)) return list;
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            list.Add(new Animal
            {
                ChipNumber = row.Cell(1).GetString(),
                Name = row.Cell(2).GetString(),
                Species = row.Cell(3).GetString(),
                // המרת מחרוזת למספר עשרוני (Double). משתמש ב-InvariantCulture כדי למנוע בעיות של פסיק/נקודה עשרונית במחשבים שונים
                Weight = double.TryParse(row.Cell(4).GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var w) ? w : 0,
                // המרת מחרוזת לתאריך
                DateOfBirth = DateTime.TryParse(row.Cell(5).GetString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : DateTime.MinValue,
                OwnerNationalId = row.Cell(6).GetString(),
                // המרה לתאריך אופציונלי (Nullable). אם נכשל מחזיר null.
                LastVaccinationDate = DateTime.TryParse(row.Cell(7).GetString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var v) ? v : null,
            });
        }
        return list;
    }

    private void WriteAnimals(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Animals");
        WriteHeader(ws, "ChipNumber", "Name", "Species", "Weight", "DateOfBirth", "OwnerNationalId", "LastVaccinationDate");
        for (int i = 0; i < Animals.Count; i++)
        {
            var a = Animals[i];
            var row = ws.Row(i + 2);
            row.Cell(1).Value = a.ChipNumber;
            row.Cell(2).Value = a.Name;
            row.Cell(3).Value = a.Species;
            // הפיכת הנתונים למחרוזת בפורמט בינלאומי (למשל תאריך yyyy-MM-dd) לשמירה אחידה
            row.Cell(4).Value = a.Weight.ToString(CultureInfo.InvariantCulture);
            row.Cell(5).Value = a.DateOfBirth.ToString("yyyy-MM-dd");
            row.Cell(6).Value = a.OwnerNationalId;
            row.Cell(7).Value = a.LastVaccinationDate?.ToString("yyyy-MM-dd") ?? "";
        }
    }

    // ---------- Visits (ביקורים) ----------

    private static List<Visit> ReadVisits(XLWorkbook wb)
    {
        var list = new List<Visit>();
        if (!wb.TryGetWorksheet("Visits", out var ws)) return list;
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            // שליפת מחרוזת התרופות (נשמרו כמופרדות בקו ניצב '|')
            var medsCsv = row.Cell(7).GetString();
            list.Add(new Visit
            {
                VisitId = row.Cell(1).GetString(),
                AnimalChipNumber = row.Cell(2).GetString(),
                VetUsername = row.Cell(3).GetString(),
                VisitDateTime = DateTime.TryParse(row.Cell(4).GetString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt) ? dt : DateTime.MinValue,
                Reason = row.Cell(5).GetString(),
                Diagnosis = row.Cell(6).GetString(),
                // אם המחרוזת ריקה מחזיר רשימה ריקה, אחרת מפצל את המחרוזת לפי '|' והופך לרשימה
                MedicationsGiven = string.IsNullOrWhiteSpace(medsCsv)
                    ? new List<string>()
                    : medsCsv.Split('|', StringSplitOptions.RemoveEmptyEntries).ToList(),
                BasePrice = decimal.TryParse(row.Cell(8).GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var bp) ? bp : 0m,
                TotalPrice = decimal.TryParse(row.Cell(9).GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var tp) ? tp : 0m,
            });
        }
        return list;
    }

    private void WriteVisits(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Visits");
        WriteHeader(ws, "VisitId", "AnimalChipNumber", "VetUsername", "VisitDateTime", "Reason", "Diagnosis", "Medications", "BasePrice", "TotalPrice");
        for (int i = 0; i < Visits.Count; i++)
        {
            var v = Visits[i];
            var row = ws.Row(i + 2);
            row.Cell(1).Value = v.VisitId;
            row.Cell(2).Value = v.AnimalChipNumber;
            row.Cell(3).Value = v.VetUsername;
            row.Cell(4).Value = v.VisitDateTime.ToString("yyyy-MM-dd HH:mm");
            row.Cell(5).Value = v.Reason;
            row.Cell(6).Value = v.Diagnosis;
            // איחוד רשימת התרופות למחרוזת אחת המופרדת ב-'|'
            row.Cell(7).Value = string.Join("|", v.MedicationsGiven);
            row.Cell(8).Value = v.BasePrice.ToString(CultureInfo.InvariantCulture);
            row.Cell(9).Value = v.TotalPrice.ToString(CultureInfo.InvariantCulture);
        }
    }

    // ---------- Medications (תרופות) ----------

    private static List<Medication> ReadMedications(XLWorkbook wb)
    {
        var list = new List<Medication>();
        if (!wb.TryGetWorksheet("Medications", out var ws)) return list;
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            list.Add(new Medication
            {
                Name = row.Cell(1).GetString(),
                Price = decimal.TryParse(row.Cell(2).GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var p) ? p : 0m,
                StockQuantity = int.TryParse(row.Cell(3).GetString(), out var s) ? s : 0,
            });
        }
        return list;
    }

    private void WriteMedications(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Medications");
        WriteHeader(ws, "Name", "Price", "StockQuantity");
        for (int i = 0; i < Medications.Count; i++)
        {
            var m = Medications[i];
            var row = ws.Row(i + 2);
            row.Cell(1).Value = m.Name;
            row.Cell(2).Value = m.Price.ToString(CultureInfo.InvariantCulture);
            row.Cell(3).Value = m.StockQuantity;
        }
    }

    // ---------- Species (סוגי חיות) ----------

    private static List<string> ReadSpecies(XLWorkbook wb)
    {
        var list = new List<string>();
        if (!wb.TryGetWorksheet("Species", out var ws)) return list;
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            var name = row.Cell(1).GetString().Trim();
            if (!string.IsNullOrEmpty(name)) list.Add(name); // מוסיף לרשימה רק אם השם לא ריק
        }
        return list;
    }

    private void WriteSpecies(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Species");
        WriteHeader(ws, "Name");
        for (int i = 0; i < Species.Count; i++)
        {
            ws.Row(i + 2).Cell(1).Value = Species[i];
        }
    }

    /// <summary>
    /// מתודת עזר המייצרת את שורת הכותרת העליונה בכל גיליון ומוסיפה לה עיצוב.
    /// </summary>
    /// <param name="ws">הגיליון הפעיל (Worksheet)</param>
    /// <param name="columns">מערך דינמי (params) של שמות העמודות</param>
    private static void WriteHeader(IXLWorksheet ws, params string[] columns)
    {
        // מעבר על כל העמודות שנשלחו
        for (int i = 0; i < columns.Length; i++)
        {
            var cell = ws.Cell(1, i + 1); // בחירת התא בשורה 1 ובעמודה המתאימה
            cell.Value = columns[i];      // הצבת שם הכותרת בתא
            cell.Style.Font.Bold = true;  // הדגשת הטקסט (Bold)
            cell.Style.Fill.BackgroundColor = XLColor.LightGray; // צביעת רקע התא באפור בהיר
        }
    }
}