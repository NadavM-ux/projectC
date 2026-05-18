using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;
using ClinicVets.Models;

namespace ClinicVets.Forms;

/// <summary>
/// מחלקה זו מייצגת את מסך ניהול הביקורים והטיפולים במערכת ClinicVets.
/// 
/// המסך אחראי על:
/// 1. הצגת רשימת כל הביקורים.
/// 2. פתיחת ביקור חדש.
/// 3. הצגת התראות חיסון שנתי.
/// 4. חישוב עלות ביקור לפי מחיר בסיס + תרופות.
/// 5. שמירת ביקור חדש למאגר הנתונים.
/// 6. עדכון תאריך חיסון אחרון במקרה שהביקור מסומן כחיסון.
/// </summary>
public class VisitsForm : Form
{
    /// <summary>
    /// אובייקט הגישה לנתונים.
    /// דרכו המחלקה ניגשת לרשימות של חיות, ביקורים ותרופות,
    /// וגם מבצעת שמירה לקובץ הנתונים.
    /// </summary>
    private readonly ExcelDataStore _store;

    /// <summary>
    /// הווטרינר המחובר כרגע למערכת.
    /// הערך הזה נשמר בתוך כל ביקור חדש בתור המטפל.
    /// </summary>
    private readonly Employee _vet;

    /// <summary>
    /// טבלת התצוגה של הביקורים במסך הראשי.
    /// הטבלה מציגה תאריך, חיה, וטרינר, סיבה, אבחנה ועלות כוללת.
    /// </summary>
    private readonly DataGridView _grid = new();

    /// <summary>
    /// תווית שמציגה התראה אם יש חיות שצריכות חיסון שנתי.
    /// </summary>
    private readonly Label _alertLabel = new();

    /// <summary>
    /// מחיר בסיס של ביקור רגיל, לפני תוספת תרופות.
    /// </summary>
    private const decimal BaseVisitPrice = 150m;

    /// <summary>
    /// מספר הימים שמגדיר תוקף חיסון שנתי.
    /// אם עברו יותר מ־365 ימים מאז החיסון האחרון, החיה תיחשב כזקוקה לחיסון.
    /// </summary>
    private const int VaccinationCadenceDays = 365;

    /// <summary>
    /// פעולה בונה של מסך הביקורים.
    /// 
    /// מקבלת את מקור הנתונים ואת הווטרינר המחובר,
    /// בונה את ממשק המשתמש ומרעננת את טבלת הביקורים.
    /// </summary>
    public VisitsForm(ExcelDataStore store, Employee vet)
    {
        // שמירת מקור הנתונים לשימוש בכל המחלקה
        _store = store;

        // שמירת פרטי הווטרינר המחובר
        _vet = vet;

        // בניית כל רכיבי המסך
        BuildUi();

        // טעינת הביקורים הקיימים לטבלה
        RefreshGrid();
    }

    /// <summary>
    /// בונה את ממשק המשתמש של מסך הביקורים.
    /// 
    /// הפעולה יוצרת:
    /// כותרת,
    /// תווית התראות,
    /// טבלת ביקורים,
    /// כפתור פתיחת ביקור חדש,
    /// וכפתור חזרה לתפריט הראשי.
    /// </summary>
    private void BuildUi()
    {
        // כותרת החלון
        Text = "Visits & Treatments";

        // גודל החלון
        Size = new Size(960, 600);

        // פתיחת החלון במרכז ביחס לחלון האב
        StartPosition = FormStartPosition.CenterParent;

        // צבע רקע כללי של המסך
        BackColor = Color.FromArgb(245, 248, 252);

        // הגדרת פונט ברירת מחדל למסך
        Font = new Font("Segoe UI", 10F);

        // יצירת כותרת ראשית למסך
        var title = new Label
        {
            Text = "🩺 Visits",
            Font = new Font("Segoe UI", 16F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            AutoSize = true,
            Location = new Point(20, 15),
        };

        // הוספת הכותרת למסך
        Controls.Add(title);

        // מיקום תווית ההתראות
        _alertLabel.Location = new Point(20, 55);

        // גודל תווית ההתראות
        _alertLabel.Size = new Size(900, 30);

        // צבע טקסט של ההתראה
        _alertLabel.ForeColor = Color.FromArgb(200, 70, 0);

        // פונט מודגש להתראה
        _alertLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);

        // הוספת התווית למסך
        Controls.Add(_alertLabel);

        // הצגת התראת חיסונים לפי מצב החיות במערכת
        ShowVaccinationAlerts();

        // מיקום טבלת הביקורים
        _grid.Location = new Point(20, 95);

        // גודל טבלת הביקורים
        _grid.Size = new Size(900, 380);

        // ביטול יצירת עמודות אוטומטית כדי להגדיר עמודות ידנית
        _grid.AutoGenerateColumns = false;

        // מניעת הוספת שורות ידנית על ידי המשתמש
        _grid.AllowUserToAddRows = false;

        // הטבלה היא לקריאה בלבד
        _grid.ReadOnly = true;

        // בחירת שורה מלאה בלחיצה
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

        // צבע רקע של הטבלה
        _grid.BackgroundColor = Color.White;

        // הסתרת עמודת כותרות השורות בצד
        _grid.RowHeadersVisible = false;

        // עמודת תאריך הביקור
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Date", DataPropertyName = nameof(Visit.VisitDateTime), Width = 130, DefaultCellStyle = { Format = "yyyy-MM-dd HH:mm" } });

        // עמודת מספר השבב של החיה
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Animal", DataPropertyName = nameof(Visit.AnimalChipNumber), Width = 100 });

        // עמודת שם המשתמש של הווטרינר
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Vet", DataPropertyName = nameof(Visit.VetUsername), Width = 100 });

        // עמודת סיבת ההגעה
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Reason", DataPropertyName = nameof(Visit.Reason), Width = 180 });

        // עמודת האבחנה
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Diagnosis", DataPropertyName = nameof(Visit.Diagnosis), Width = 180 });

        // עמודת המחיר הכולל של הביקור
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Total (₪)", DataPropertyName = nameof(Visit.TotalPrice), Width = 100 });

        // הוספת הטבלה למסך
        Controls.Add(_grid);

        // כפתור לפתיחת ביקור חדש
        var addBtn = new Button
        {
            Text = "➕ Open New Visit",
            Location = new Point(20, 495),
            Size = new Size(180, 38),
            BackColor = Color.FromArgb(40, 160, 90),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
        };

        // הסרת מסגרת הכפתור
        addBtn.FlatAppearance.BorderSize = 0;

        // קישור אירוע הלחיצה לפונקציה שפותחת ביקור חדש
        addBtn.Click += OnNewVisit;

        // הוספת הכפתור למסך
        Controls.Add(addBtn);

        // כפתור חזרה לתפריט הראשי
        var backBtn = new Button
        {
            Text = "← Main Menu",
            Location = new Point(220, 495),
            Size = new Size(150, 38),
            BackColor = Color.FromArgb(120, 130, 150),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
        };

        // הסרת מסגרת הכפתור
        backBtn.FlatAppearance.BorderSize = 0;

        // סגירת המסך הנוכחי בעת לחיצה
        backBtn.Click += (_, _) => Close();

        // הוספת הכפתור למסך
        Controls.Add(backBtn);
    }

    /// <summary>
    /// מרעננת את טבלת הביקורים.
    /// 
    /// הפעולה מנקה את מקור הנתונים הקודם של הטבלה,
    /// ולאחר מכן טוענת את כל הביקורים מחדש,
    /// כשהם מסודרים מהביקור החדש ביותר לישן ביותר.
    /// </summary>
    private void RefreshGrid()
    {
        // ניקוי מקור הנתונים כדי לאפשר רענון מלא
        _grid.DataSource = null;

        // טעינת הביקורים מחדש, ממוינים לפי תאריך בסדר יורד
        _grid.DataSource = _store.Visits.OrderByDescending(v => v.VisitDateTime).ToList();
    }

    /// <summary>
    /// מציגה התראה לגבי חיות שצריכות חיסון שנתי.
    /// 
    /// חיה תיחשב כזקוקה לחיסון אם:
    /// 1. אין לה תאריך חיסון אחרון.
    /// או
    /// 2. עברו יותר מ־365 ימים מאז החיסון האחרון שלה.
    /// </summary>
    private void ShowVaccinationAlerts()
    {
        // שליפת כל החיות שצריכות חיסון שנתי
        var due = _store.Animals
            .Where(a => a.LastVaccinationDate == null
                        || (DateTime.Now - a.LastVaccinationDate.Value).TotalDays > VaccinationCadenceDays)
            .ToList();

        // אם אין חיות שצריכות חיסון, מוצגת הודעה חיובית.
        // אחרת, מוצגת אזהרה עם מספר החיות ושמות עד 5 חיות ראשונות.
        _alertLabel.Text = due.Count == 0
            ? "✅ All animals are up to date on annual vaccinations."
            : $"⚠ {due.Count} animal(s) need an annual vaccination: " + string.Join(", ", due.Take(5).Select(a => a.Name));
    }

    /// <summary>
    /// מופעלת כאשר המשתמש לוחץ על הכפתור Open New Visit.
    /// 
    /// הפעולה:
    /// 1. בודקת שיש חיות במערכת.
    /// 2. פותחת חלון יצירת ביקור חדש.
    /// 3. אם הביקור נשמר בהצלחה, מוסיפה אותו לרשימת הביקורים.
    /// 4. אם סומן חיסון, מעדכנת את החיה המתאימה.
    /// 5. שומרת את הנתונים ומרעננת את המסך.
    /// </summary>
    private void OnNewVisit(object? sender, EventArgs e)
    {
        // אם אין חיות במערכת, אין אפשרות לפתוח ביקור
        if (_store.Animals.Count == 0)
        {
            MessageBox.Show("There are no registered animals.", "No animals", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // יצירת חלון פתיחת ביקור חדש
        using var dlg = new NewVisitDialog(_store, _vet, BaseVisitPrice);

        // אם המשתמש לחץ Save והביקור נוצר בהצלחה
        if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Result != null)
        {
            // הוספת הביקור החדש לרשימת הביקורים
            _store.Visits.Add(dlg.Result);


            // מעבר על כל התרופות שניתנו במהלך הביקור
            foreach (var medName in dlg.Result.MedicationsGiven)
            {
                // חיפוש התרופה במערכת לפי שם
                var medication = _store.Medications
                    .FirstOrDefault(m => m.Name == medName);

                // אם התרופה נמצאה ויש מלאי גדול מ־0
                if (medication != null && medication.StockQuantity > 0)
                {
                    // הורדת יחידה אחת מהמלאי
                    medication.StockQuantity--;
                }
            }

            // אם במהלך הביקור עודכן תאריך חיסון של חיה
            if (dlg.UpdatedAnimal != null)
            {
                // חיפוש החיה המקורית לפי מספר שבב
                var idx = _store.Animals.FindIndex(a => a.ChipNumber == dlg.UpdatedAnimal.ChipNumber);

                // אם החיה נמצאה, מחליפים אותה בגרסה המעודכנת
                if (idx >= 0) _store.Animals[idx] = dlg.UpdatedAnimal;
            }

            // שמירת כל הנתונים למאגר
            _store.Save();

            // רענון התראת חיסונים
            ShowVaccinationAlerts();

            // רענון טבלת הביקורים
            RefreshGrid();
        }
    }

    /// <summary>
    /// מחלקה פנימית שמייצגת את חלון פתיחת ביקור חדש.
    /// 
    /// החלון מאפשר:
    /// 1. בחירת חיה.
    /// 2. בחירת תאריך ושעה.
    /// 3. הזנת סיבת ביקור.
    /// 4. הזנת אבחנה.
    /// 5. בחירת תרופות שניתנו.
    /// 6. סימון שהביקור כולל חיסון.
    /// 7. חישוב מחיר כולל.
    /// </summary>
    private class NewVisitDialog : Form
    {
        /// <summary>
        /// מקור הנתונים של המערכת.
        /// משמש לגישה לחיות ולתרופות.
        /// </summary>
        private readonly ExcelDataStore _store;

        /// <summary>
        /// הווטרינר שמבצע את הביקור.
        /// </summary>
        private readonly Employee _vet;

        /// <summary>
        /// מחיר בסיס של ביקור.
        /// </summary>
        private readonly decimal _basePrice;

        /// <summary>
        /// רשימת בחירה של החיות במערכת.
        /// </summary>
        private readonly ComboBox _animal = new();

        /// <summary>
        /// רכיב לבחירת תאריך ושעת הביקור.
        /// </summary>
        private readonly DateTimePicker _date = new();

        /// <summary>
        /// שדה טקסט לסיבת ההגעה.
        /// </summary>
        private readonly TextBox _reason = new();

        /// <summary>
        /// שדה טקסט לאבחנה רפואית.
        /// </summary>
        private readonly TextBox _diagnosis = new();

        /// <summary>
        /// רשימת תרופות עם אפשרות סימון.
        /// כל תרופה מוצגת עם השם והמחיר שלה.
        /// </summary>
        private readonly CheckedListBox _meds = new();

        /// <summary>
        /// תווית המציגה את המחיר הכולל של הביקור.
        /// </summary>
        private readonly Label _totalLabel = new();

        /// <summary>
        /// סימון האם הביקור הנוכחי הוא חיסון של היום.
        /// אם מסומן, תאריך החיסון האחרון של החיה יעודכן.
        /// </summary>
        private readonly CheckBox _markVaccinated = new();

        /// <summary>
        /// תוצאת החלון: הביקור החדש שנוצר.
        /// אם המשתמש ביטל את הפעולה, הערך יישאר null.
        /// </summary>
        public Visit? Result { get; private set; }

        /// <summary>
        /// חיה מעודכנת במקרה שסומן שהביקור הוא חיסון.
        /// משמש כדי לעדכן את LastVaccinationDate במסך הראשי.
        /// </summary>
        public Animal? UpdatedAnimal { get; private set; }

        /// <summary>
        /// פעולה בונה של חלון יצירת ביקור חדש.
        /// 
        /// מקבלת:
        /// store - מקור הנתונים.
        /// vet - הווטרינר המטפל.
        /// basePrice - מחיר בסיס של ביקור.
        /// </summary>
        public NewVisitDialog(ExcelDataStore store, Employee vet, decimal basePrice)
        {
            // שמירת מקור הנתונים
            _store = store;

            // שמירת הווטרינר המטפל
            _vet = vet;

            // שמירת מחיר הבסיס
            _basePrice = basePrice;

            // כותרת החלון
            Text = "Open New Visit";

            // גודל החלון
            Size = new Size(560, 640);

            // פתיחת החלון במרכז
            StartPosition = FormStartPosition.CenterParent;

            // צבע רקע
            BackColor = Color.FromArgb(245, 248, 252);

            // פונט כללי
            Font = new Font("Segoe UI", 10F);

            // חלון דיאלוג קבוע
            FormBorderStyle = FormBorderStyle.FixedDialog;

            // ביטול כפתור הגדלה
            MaximizeBox = false;

            // תווית בחירת חיה
            var lblAnimal = new Label { Text = "Animal:", Location = new Point(20, 23), AutoSize = true };

            // מיקום וגודל רשימת החיות
            _animal.Location = new Point(150, 20);
            _animal.Size = new Size(370, 25);

            // המשתמש יכול לבחור רק מתוך הרשימה ולא להקליד חיה ידנית
            _animal.DropDownStyle = ComboBoxStyle.DropDownList;

            // טעינת כל החיות מתוך המאגר אל הרשימה
            foreach (var a in _store.Animals)
            {
                _animal.Items.Add($"{a.Name} (chip {a.ChipNumber}) — {a.Species}");
            }

            // בחירת החיה הראשונה כברירת מחדל
            _animal.SelectedIndex = 0;

            // הוספת התווית ורשימת החיות לחלון
            Controls.Add(lblAnimal); Controls.Add(_animal);

            // תווית תאריך ושעה
            var lblDate = new Label { Text = "Date / time:", Location = new Point(20, 63), AutoSize = true };

            // מיקום וגודל בוחר התאריך
            _date.Location = new Point(150, 60);
            _date.Size = new Size(370, 25);

            // הגדרת פורמט מותאם אישית
            _date.Format = DateTimePickerFormat.Custom;

            // פורמט התאריך והשעה שיוצג למשתמש
            _date.CustomFormat = "yyyy-MM-dd HH:mm";

            // ערך ברירת מחדל: התאריך והשעה הנוכחיים
            _date.Value = DateTime.Now;

            // הוספת רכיבי התאריך לחלון
            Controls.Add(lblDate); Controls.Add(_date);

            // תווית סיבת ביקור
            var lblReason = new Label { Text = "Reason for visit:", Location = new Point(20, 103), AutoSize = true };

            // מיקום וגודל שדה סיבת הביקור
            _reason.Location = new Point(150, 100);
            _reason.Size = new Size(370, 25);

            // הוספת סיבת הביקור לחלון
            Controls.Add(lblReason); Controls.Add(_reason);

            // תווית אבחנה
            var lblDiag = new Label { Text = "Diagnosis:", Location = new Point(20, 143), AutoSize = true };

            // מיקום וגודל שדה האבחנה
            _diagnosis.Location = new Point(150, 140);
            _diagnosis.Size = new Size(370, 60);

            // מאפשר כתיבה בכמה שורות
            _diagnosis.Multiline = true;

            // הוספת שדה האבחנה לחלון
            Controls.Add(lblDiag); Controls.Add(_diagnosis);

            // כותרת אזור התרופות
            var lblMeds = new Label { Text = "Medications administered:", Location = new Point(20, 213), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold) };

            // מיקום וגודל רשימת התרופות
            _meds.Location = new Point(20, 240);
            _meds.Size = new Size(500, 180);

            // מאפשר סימון תרופה בלחיצה אחת
            _meds.CheckOnClick = true;

            // מעבר על כל התרופות במערכת
            foreach (var m in _store.Medications)
            {
                // אם אין מלאי לתרופה
                if (m.StockQuantity <= 0)
                {
                    // מציגים שהתרופה אזלה מהמלאי
                    _meds.Items.Add($"{m.Name} — OUT OF STOCK");
                }
                else
                {
                    // תרופה זמינה עם מחיר ומלאי
                    _meds.Items.Add($"{m.Name} — ₪{m.Price:0.##} (Stock: {m.StockQuantity})");
                }
            }

            _meds.ItemCheck += (sender, e) =>
            {
                // הטקסט של התרופה שנבחרה
                var itemText = _meds.Items[e.Index]?.ToString() ?? "";

                // אם התרופה אזלה מהמלאי
                if (itemText.Contains("OUT OF STOCK"))
                {
                    // ביטול הסימון
                    e.NewValue = CheckState.Unchecked;

                    // הודעה למשתמש
                    MessageBox.Show(
                        "This medication is out of stock.",
                        "Out of stock",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // עדכון המחיר הכולל אחרי הסימון
                BeginInvoke(new Action(UpdateTotal));
            };

            // הוספת התרופות לחלון
            Controls.Add(lblMeds); Controls.Add(_meds);

            // טקסט של סימון חיסון
            _markVaccinated.Text = "Mark this visit as today's vaccination";

            // מיקום סימון החיסון
            _markVaccinated.Location = new Point(20, 430);

            // התאמת גודל אוטומטית לטקסט
            _markVaccinated.AutoSize = true;

            // הוספת סימון החיסון לחלון
            Controls.Add(_markVaccinated);

            // הצגת מחיר התחלתי: מחיר בסיס בלבד
            _totalLabel.Text = $"Total: ₪{_basePrice:0.##} (base)";

            // מיקום תווית המחיר
            _totalLabel.Location = new Point(20, 470);

            // התאמת גודל אוטומטית
            _totalLabel.AutoSize = true;

            // פונט מודגש למחיר
            _totalLabel.Font = new Font("Segoe UI", 12F, FontStyle.Bold);

            // צבע המחיר
            _totalLabel.ForeColor = Color.FromArgb(40, 90, 150);

            // הוספת המחיר לחלון
            Controls.Add(_totalLabel);

            // כפתור שמירת ביקור
            var save = new Button
            {
                Text = "Save Visit",
                Location = new Point(280, 510),
                Size = new Size(110, 36),
                BackColor = Color.FromArgb(40, 160, 90),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            };

            // הסרת מסגרת מכפתור השמירה
            save.FlatAppearance.BorderSize = 0;

            // קישור לחיצה לפעולת שמירה
            save.Click += OnSave;

            // כפתור ביטול
            var cancel = new Button
            {
                Text = "Cancel",
                Location = new Point(400, 510),
                Size = new Size(110, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(220, 220, 220),
            };

            // הסרת מסגרת מכפתור הביטול
            cancel.FlatAppearance.BorderSize = 0;

            // סגירת החלון ללא שמירה
            cancel.Click += (_, _) => Close();

            // הוספת כפתורי השמירה והביטול
            Controls.Add(save); Controls.Add(cancel);
        }

        /// <summary>
        /// מחשבת מחדש את העלות הכוללת של הביקור.
        /// 
        /// החישוב:
        /// מחיר כולל = מחיר בסיס + סכום מחירי התרופות שסומנו.
        /// 
        /// הערה:
        /// המחיר של התרופה נשלף מתוך הטקסט שמוצג ברשימת התרופות.
        /// </summary>
        private void UpdateTotal()
        {
            // התחלה ממחיר הבסיס של ביקור
            decimal total = _basePrice;

            // מעבר על כל התרופות שסומנו
            foreach (var item in _meds.CheckedItems)
            {
                // המרת הפריט לטקסט
                var text = item?.ToString() ?? "";

                // פיצול לפי סימן השקל
                var parts = text.Split('₪');

                // אם הצלחנו לקבל חלק מחיר
                if (parts.Length >= 2)
                {
                    // ניקוי כל מה שאחרי הרווח הראשון
                    var pricePart = parts[1].Split(' ')[0];

                    // ניסיון להמיר למחיר
                    if (decimal.TryParse(pricePart, out var p))
                    {
                        // הוספת המחיר לסכום הכולל
                        total += p;
                    }
                }
            }

            // עדכון תצוגת המחיר למשתמש
            _totalLabel.Text = $"Total: ₪{total:0.##} (base ₪{_basePrice:0.##} + meds ₪{total - _basePrice:0.##})";
        }

        /// <summary>
        /// מופעלת כאשר המשתמש לוחץ על Save Visit.
        /// 
        /// הפעולה:
        /// 1. בודקת ששדות חובה מולאו.
        /// 2. מוציאה את מספר השבב מתוך טקסט החיה שנבחרה.
        /// 3. אוספת את התרופות שסומנו.
        /// 4. מחשבת את עלות התרופות.
        /// 5. יוצרת אובייקט Visit חדש.
        /// 6. אם סומן חיסון, יוצרת Animal מעודכן עם תאריך חיסון חדש.
        /// 7. סוגרת את החלון עם DialogResult.OK.
        /// </summary>
        private void OnSave(object? sender, EventArgs e)
        {
            // בדיקת חובה: סיבת ביקור
            if (string.IsNullOrWhiteSpace(_reason.Text))
            { Tell("Reason is required."); return; }

            // בדיקת חובה: אבחנה
            if (string.IsNullOrWhiteSpace(_diagnosis.Text))
            { Tell("Diagnosis is required."); return; }

            // חייבת להיות לפחות אות אחת בסיבת הביקור
            if (!_reason.Text.Any(char.IsLetter))
            {
                Tell("Reason must contain letters.");
                return;
            }

            // חייבת להיות לפחות אות אחת באבחנה
            if (!_diagnosis.Text.Any(char.IsLetter))
            {
                Tell("Diagnosis must contain letters.");
                return;
            }

            // בדיקת חובה: בחירת חיה
            if (_animal.SelectedItem == null) { Tell("Choose an animal."); return; }

            // קבלת הטקסט של החיה שנבחרה
            var animalText = _animal.SelectedItem.ToString() ?? "";

            // מציאת תחילת מספר השבב מתוך הטקסט
            var chipStart = animalText.IndexOf("chip ", StringComparison.Ordinal) + 5;

            // מציאת סוף מספר השבב
            var chipEnd = animalText.IndexOf(')', chipStart);

            // חילוץ מספר השבב
            var chip = animalText.Substring(chipStart, chipEnd - chipStart);

            // רשימת שמות התרופות שניתנו
            var meds = new List<string>();

            // סכום מחירי התרופות
            decimal medTotal = 0m;

            // מעבר על כל התרופות שסומנו
            foreach (var item in _meds.CheckedItems)
            {
                // המרת התרופה לטקסט
                var text = item?.ToString() ?? "";

                // מציאת המפריד בין שם התרופה למחיר
                var dashIdx = text.IndexOf(" — ", StringComparison.Ordinal);

                // חילוץ שם התרופה
                var name = dashIdx > 0 ? text[..dashIdx] : text;

                // הוספת שם התרופה לרשימת התרופות שניתנו
                meds.Add(name);

                // חילוץ מחיר התרופה מתוך הטקסט
                var priceStr = text.Split('₪').LastOrDefault() ?? "0";

                // אם המחיר תקין, מוסיפים אותו לסכום התרופות
                if (decimal.TryParse(priceStr, out var p)) medTotal += p;
            }

            // יצירת אובייקט ביקור חדש
            Result = new Visit
            {
                // מזהה קצר וייחודי לביקור
                VisitId = Guid.NewGuid().ToString("N")[..8],

                // מספר השבב של החיה
                AnimalChipNumber = chip,

                // שם המשתמש של הווטרינר המטפל
                VetUsername = _vet.Username,

                // תאריך ושעת הביקור
                VisitDateTime = _date.Value,

                // סיבת הביקור לאחר Trim
                Reason = _reason.Text.Trim(),

                // אבחנה לאחר Trim
                Diagnosis = _diagnosis.Text.Trim(),

                // רשימת התרופות שניתנו
                MedicationsGiven = meds,

                // מחיר בסיס
                BasePrice = _basePrice,

                // מחיר כולל
                TotalPrice = _basePrice + medTotal,
            };

            // אם הביקור מסומן כחיסון של היום
            if (_markVaccinated.Checked)
            {
                // חיפוש החיה לפי מספר שבב
                var animal = _store.Animals.FirstOrDefault(a => a.ChipNumber == chip);

                // אם החיה נמצאה
                if (animal != null)
                {
                    // יצירת עותק מעודכן של החיה עם תאריך חיסון חדש
                    UpdatedAnimal = new Animal
                    {
                        ChipNumber = animal.ChipNumber,
                        Name = animal.Name,
                        Species = animal.Species,
                        Weight = animal.Weight,
                        DateOfBirth = animal.DateOfBirth,
                        OwnerNationalId = animal.OwnerNationalId,

                        // עדכון תאריך החיסון האחרון לתאריך הביקור
                        LastVaccinationDate = _date.Value.Date,
                    };
                }
            }

            // סימון שהחלון הסתיים בהצלחה
            DialogResult = DialogResult.OK;

            // סגירת החלון
            Close();
        }

        /// <summary>
        /// מציגה הודעת שגיאת ולידציה למשתמש.
        /// 
        /// msg - ההודעה שתוצג.
        /// </summary>
        private static void Tell(string msg) =>
            MessageBox.Show(msg, "Validation error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}