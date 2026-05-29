using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;
using ClinicVets.Models;

namespace ClinicVets.Forms;

/// <summary>
/// מסך ניהול מלאי התרופות במערכת ClinicVets.
///
/// המסך מאפשר:
/// 1. צפייה בכל התרופות הקיימות.
/// 2. הוספת תרופה חדשה.
/// 3. מחיקת תרופה קיימת.
/// 4. שמירת השינויים למאגר הנתונים.
/// </summary>
public class MedicationsForm : Form
{
    /// <summary>
    /// מקור הנתונים הראשי של המערכת.
    ///
    /// דרכו מתבצעת גישה:
    /// - לרשימת התרופות
    /// - לשמירת הנתונים
    /// </summary>
    private readonly ExcelDataStore _store;

    /// <summary>
    /// טבלת התצוגה של התרופות.
    ///
    /// מציגה:
    /// - שם תרופה
    /// - מחיר
    /// - כמות מלאי
    /// </summary>
    private readonly DataGridView _grid = new();



    // תווית חיפוש תרופה
    private readonly Label _searchLabel = new();

    // תיבת חיפוש תרופה
    private readonly TextBox _searchBox = new();

    // כפתורים של המסך
    // כפתורים של המסך
    private Button _addBtn = null!;
    private Button _deleteBtn = null!;
    private Button _editBtn = null!;
    private Button _backBtn = null!;

    /// <summary>
    /// פעולה בונה של מסך ניהול התרופות.
    ///
    /// מקבלת את מקור הנתונים של המערכת,
    /// בונה את ממשק המשתמש,
    /// ומרעננת את טבלת התרופות.
    /// </summary>
    public MedicationsForm(ExcelDataStore store)
    {
        // שמירת מקור הנתונים לשימוש במחלקה
        _store = store;

        // בניית כל רכיבי ה־UI של המסך
        BuildUi();

        // טעינת התרופות לטבלה
        RefreshGrid();
    }

    /// <summary>
    /// בונה את כל ממשק המשתמש של מסך התרופות.
    ///
    /// הפעולה יוצרת:
    /// - כותרת
    /// - טבלת תרופות
    /// - כפתור הוספה
    /// - כפתור מחיקה
    /// - כפתור חזרה
    /// </summary>
    private void BuildUi()
    {
        // כותרת החלון
        Text = "Medication Inventory";

        // גודל החלון נשאר כמו שהיה
        Size = new Size(700, 520);

        // מאפשר להגדיל את החלון
        FormBorderStyle = FormBorderStyle.Sizable;

        // מאפשר כפתור הגדלה
        MaximizeBox = true;

        // לא נותן להקטין את החלון מתחת לגודל התקין
        MinimumSize = new Size(700, 520);

        // פתיחת החלון במרכז
        StartPosition = FormStartPosition.CenterParent;

        // צבע רקע כללי
        BackColor = Color.FromArgb(245, 248, 252);

        // פונט ברירת מחדל
        Font = new Font("Segoe UI", 10F);

        // יצירת כותרת ראשית למסך
        var title = new Label
        {
            Text = "💊 Medications",

            // פונט גדול ומודגש
            Font = new Font("Segoe UI", 16F, FontStyle.Bold),

            // צבע הכותרת
            ForeColor = Color.FromArgb(40, 90, 150),

            // התאמת גודל אוטומטית
            AutoSize = true,

            // מיקום על המסך
            Location = new Point(20, 15),
        };

        // הוספת הכותרת למסך
        Controls.Add(title);

        // תווית חיפוש תרופה
        _searchLabel.Text = "🔍 Search:";
        _searchLabel.AutoSize = true;
        _searchLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
        _searchLabel.ForeColor = Color.FromArgb(40, 90, 150);
        _searchLabel.Location = new Point(390, 28);

        // תיבת חיפוש תרופה
        _searchBox.Location = new Point(480, 25);
        _searchBox.Size = new Size(180, 28);
        _searchBox.Font = new Font("Segoe UI", 10F);
        _searchBox.PlaceholderText = "Search medication...";
        _searchBox.BorderStyle = BorderStyle.FixedSingle;

        // בכל הקלדה מרעננים את הטבלה לפי החיפוש
        _searchBox.TextChanged += (_, _) => RefreshGrid(_searchBox.Text);

        // הוספת החיפוש למסך
        Controls.Add(_searchLabel);
        Controls.Add(_searchBox);

        // מיקום טבלת התרופות
        _grid.Location = new Point(20, 65);

        // גודל הטבלה - מוגדל קצת בלי לשנות את גודל החלון
        _grid.Size = new Size(640, 355);

        // ביטול יצירת עמודות אוטומטית
        _grid.AutoGenerateColumns = false;

        // מניעת הוספת שורות ידנית
        _grid.AllowUserToAddRows = false;

        // מצב קריאה בלבד
        _grid.ReadOnly = true;

        // בחירת שורה מלאה
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

        // מאפשר בחירת שורה אחת בלבד
        _grid.MultiSelect = false;

        // צבע רקע הטבלה
        _grid.BackgroundColor = Color.White;

        // הסתרת כותרות שורות
        _grid.RowHeadersVisible = false;

        // העמודות יתפרסו על כל רוחב הטבלה
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        // אם יהיו הרבה תרופות, תופיע גלילה אנכית
        _grid.ScrollBars = ScrollBars.Vertical;

        // מניעת שינוי גובה שורות ידני
        _grid.AllowUserToResizeRows = false;

        /// <summary>
        /// עמודת שם התרופה.
        /// 
        /// DataPropertyName מחבר את העמודה
        /// לשדה Name מתוך אובייקט Medication.
        /// </summary>
        _grid.Columns.Add(
            new DataGridViewTextBoxColumn
            {
                HeaderText = "Name",
                DataPropertyName = nameof(Medication.Name),
                Width = 280
            });

        /// <summary>
        /// עמודת מחיר התרופה.
        /// </summary>
        _grid.Columns.Add(
            new DataGridViewTextBoxColumn
            {
                HeaderText = "Price (₪)",
                DataPropertyName = nameof(Medication.Price),
                Width = 130
            });

        /// <summary>
        /// עמודת כמות מלאי.
        /// </summary>
        _grid.Columns.Add(
            new DataGridViewTextBoxColumn
            {
                HeaderText = "Stock",
                DataPropertyName = nameof(Medication.StockQuantity),
                Width = 130
            });

        // הוספת הטבלה למסך
        Controls.Add(_grid);

        /// <summary>
        /// כפתור הוספת תרופה חדשה.
        /// </summary>
        _addBtn = MakeButton(
            "➕ Add",
            new Point(20, 430),
            Color.FromArgb(40, 160, 90));

        // בעת לחיצה → פתיחת חלון הוספת תרופה
        _addBtn.Click += OnAdd;


        /// <summary>
        /// כפתור מחיקת תרופה.
        /// </summary>
        _deleteBtn = MakeButton(
            "🗑 Delete",
            new Point(130, 430),
            Color.FromArgb(200, 70, 70));

        // בעת לחיצה → מחיקת התרופה הנבחרת
        _deleteBtn.Click += OnDelete;

        /// <summary>
        /// כפתור חזרה לתפריט הראשי.
        /// </summary>
        _backBtn = MakeButton(
            "← Main Menu",
            new Point(370, 430),
            Color.FromArgb(120, 130, 150));

        // שינוי גודל הכפתור
        _backBtn.Size = new Size(140, 36);

        // סגירת המסך בעת לחיצה
        _backBtn.Click += (_, _) => Close();


        /// <summary>
        /// כפתור עריכת תרופה.
        /// </summary>
        _editBtn = MakeButton(
            "✏ Edit",
            new Point(240, 430),
            Color.FromArgb(52, 152, 219));

        // בעת לחיצה → עריכת התרופה הנבחרת
        _editBtn.Click += OnEdit;





        Controls.Add(_addBtn);
        Controls.Add(_deleteBtn);
        Controls.Add(_editBtn);
        Controls.Add(_backBtn);

        // התאמת המסך כאשר מגדילים או מקטינים אותו
        Resize += MedicationsForm_Resize;

        // הפעלה ראשונית כדי שהכול יהיה במקום כבר בפתיחה
        MedicationsForm_Resize(this, EventArgs.Empty);
    }





    // מסדר את מסך התרופות לפי גודל החלון
    private void MedicationsForm_Resize(object? sender, EventArgs e)
    {
        int margin = 20;

        // מיקום וגודל של החיפוש
        _searchBox.Size = new Size(210, 28);

        _searchBox.Location = new Point(
            ClientSize.Width - _searchBox.Width - margin,
            25
        );

        _searchLabel.Location = new Point(
            _searchBox.Left - _searchLabel.Width - 10,
            28
        );

        // הטבלה מתרחבת עם החלון
        _grid.Location = new Point(margin, 65);

        _grid.Size = new Size(
            ClientSize.Width - (margin * 2),
            ClientSize.Height - 150
        );

        // כפתורים בתחתית
        int buttonY = ClientSize.Height - 60;

        // מצב חלון גדול
        if (ClientSize.Width > 1000)
        {
            _grid.DefaultCellStyle.Font = new Font("Segoe UI", 13F);
            _grid.ColumnHeadersDefaultCellStyle.Font =
                new Font("Segoe UI", 14F, FontStyle.Bold);

            _grid.ColumnHeadersHeight = 50;

            foreach (DataGridViewRow row in _grid.Rows)
            {
                row.Height = 45;
            }

            _addBtn.Size = new Size(130, 45);
            _deleteBtn.Size = new Size(130, 45);
            _editBtn.Size = new Size(130, 45);
            _backBtn.Size = new Size(170, 45);

            _addBtn.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            _deleteBtn.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            _editBtn.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            _backBtn.Font = new Font("Segoe UI", 11F, FontStyle.Bold);

            _searchBox.Font = new Font("Segoe UI", 11F);
            _searchLabel.Font = new Font("Segoe UI", 11F, FontStyle.Bold);

            // רווח מסודר בין הכפתורים
            int gap = 25;

            _addBtn.Location = new Point(margin, buttonY);
            _deleteBtn.Location = new Point(_addBtn.Right + gap, buttonY);
            _editBtn.Location = new Point(_deleteBtn.Right + gap, buttonY);
            _backBtn.Location = new Point(_editBtn.Right + gap, buttonY);
        }
        else
        {
            _grid.DefaultCellStyle.Font = new Font("Segoe UI", 10F);
            _grid.ColumnHeadersDefaultCellStyle.Font =
                new Font("Segoe UI", 10F, FontStyle.Regular);

            _grid.ColumnHeadersHeight = 32;

            foreach (DataGridViewRow row in _grid.Rows)
            {
                row.Height = 32;
            }

            _addBtn.Size = new Size(100, 36);
            _deleteBtn.Size = new Size(100, 36);
            _editBtn.Size = new Size(100, 36);
            _backBtn.Size = new Size(140, 36);

            _addBtn.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            _deleteBtn.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            _editBtn.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            _backBtn.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

            _searchBox.Font = new Font("Segoe UI", 10F);
            _searchLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);

            // רווח רגיל במסך קטן
            int gap = 10;

            _addBtn.Location = new Point(margin, buttonY);
            _deleteBtn.Location = new Point(_addBtn.Right + gap, buttonY);
            _editBtn.Location = new Point(_deleteBtn.Right + gap, buttonY);
            _backBtn.Location = new Point(_editBtn.Right + 50, buttonY);
        }
    }





    /// <summary>
    /// פונקציית עזר ליצירת כפתורים מעוצבים.
    ///
    /// הפונקציה מקבלת:
    /// text - טקסט הכפתור.
    /// location - מיקום הכפתור.
    /// back - צבע הרקע.
    ///
    /// ומחזירה Button מוכן לשימוש.
    /// </summary>
    private static Button MakeButton(string text, Point location, Color back)
    {
        // יצירת כפתור חדש
        var b = new Button
        {
            // טקסט שיוצג
            Text = text,

            // מיקום הכפתור
            Location = location,

            // גודל הכפתור
            Size = new Size(100, 36),

            // צבע רקע
            BackColor = back,

            // צבע טקסט
            ForeColor = Color.White,

            // עיצוב שטוח
            FlatStyle = FlatStyle.Flat,

            // פונט מודגש
            Font = new Font("Segoe UI", 9F, FontStyle.Bold),
        };

        // הסרת מסגרת הכפתור
        b.FlatAppearance.BorderSize = 0;

        // החזרת הכפתור שנוצר
        return b;
    }

    /// <summary>
    /// מרעננת את טבלת התרופות.
    ///
    /// הפעולה:
    /// 1. מנקה את הנתונים הישנים.
    /// 2. טוענת מחדש את רשימת התרופות.
    /// </summary>
    private void RefreshGrid(string filter = "")
    {
        // ניקוי מקור הנתונים הישן
        _grid.DataSource = null;

        // טקסט החיפוש
        string searchText = filter.Trim().ToLower();

        // אם אין חיפוש, מציגים את כל התרופות
        if (string.IsNullOrWhiteSpace(searchText))
        {
            _grid.DataSource = _store.Medications.ToList();
        }
        else
        {
            // אם יש חיפוש, מציגים רק תרופות שהשם שלהן מכיל את הטקסט
            _grid.DataSource = _store.Medications
                .Where(m => m.Name.ToLower().Contains(searchText))
                .ToList();
        }

        // אחרי רענון הנתונים מחזירים את העיצוב לפי גודל החלון
        MedicationsForm_Resize(this, EventArgs.Empty);
    }

    /// <summary>
    /// מופעלת כאשר המשתמש לוחץ על Add.
    ///
    /// הפעולה:
    /// 1. פותחת חלון הוספת תרופה.
    /// 2. בודקת שאין תרופה כפולה.
    /// 3. מוסיפה את התרופה למערכת.
    /// 4. שומרת את הנתונים.
    /// 5. מרעננת את הטבלה.
    /// </summary>
    private void OnAdd(object? sender, EventArgs e)
    {
        // פתיחת חלון הוספת תרופה
        using var dlg = new AddMedicationDialog();

        // אם המשתמש לחץ Save ונוצרה תרופה תקינה
        if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Result != null)
        {
            /// <summary>
            /// בדיקת duplicate:
            /// האם כבר קיימת תרופה עם אותו שם.
            ///
            /// StringComparison.OrdinalIgnoreCase
            /// אומר שמתעלמים מאותיות גדולות/קטנות.
            /// </summary>
            if (_store.Medications.Any(
                    m => m.Name.Equals(
                        dlg.Result.Name,
                        StringComparison.OrdinalIgnoreCase)))
            {
                // הודעת שגיאה במקרה של duplicate
                MessageBox.Show(
                    "A medication with this name already exists.",
                    "Duplicate",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            // הוספת התרופה לרשימת התרופות
            _store.Medications.Add(dlg.Result);

            // שמירת הנתונים
            _store.Save();

            // רענון הטבלה
            RefreshGrid(_searchBox.Text);
        }
    }

    /// <summary>
    /// מופעלת כאשר המשתמש לוחץ על Delete.
    ///
    /// הפעולה:
    /// 1. בודקת שנבחרה תרופה.
    /// 2. מבקשת אישור מחיקה.
    /// 3. מוחקת את התרופה.
    /// 4. שומרת את הנתונים.
    /// 5. מרעננת את הטבלה.
    /// </summary>
    private void OnDelete(object? sender, EventArgs e)
    {
        // אם לא נבחרה שורה או שהשורה אינה Medication
        if (_grid.CurrentRow?.DataBoundItem is not Medication m)
            return;

        // חלון אישור מחיקה
        var confirm = MessageBox.Show(
            $"Delete '{m.Name}'?",
            "Confirm",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        // אם המשתמש לא אישר
        if (confirm != DialogResult.Yes)
            return;

        /// <summary>
        /// מחיקת כל התרופות עם אותו שם.
        ///
        /// RemoveAll מחזיר את כל האיברים
        /// שעומדים בתנאי.
        /// </summary>
        _store.Medications.RemoveAll(
            med => med.Name == m.Name);

        // שמירת הנתונים
        _store.Save();

        // רענון הטבלה
        RefreshGrid(_searchBox.Text);
    }



    /// <summary>
    /// מופעלת כאשר המשתמש לוחץ על Edit.
    ///
    /// הפעולה:
    /// 1. בודקת שנבחרה תרופה.
    /// 2. פותחת חלון עריכה.
    /// 3. שומרת את השינויים.
    /// 4. מרעננת את הטבלה.
    /// </summary>
    private void OnEdit(object? sender, EventArgs e)
    {
        // אם לא נבחרה תרופה
        if (_grid.CurrentRow?.DataBoundItem is not Medication medication)
        {
            MessageBox.Show(
                "Please select a medication first.",
                "No selection",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);

            return;
        }

        // פתיחת חלון העריכה
        using var dlg = new EditMedicationDialog(medication);

        // אם המשתמש לחץ Save
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            // שמירת הנתונים
            _store.Save();

            // רענון הטבלה
            RefreshGrid(_searchBox.Text);
        }
    }

    /// <summary>
    /// חלון פנימי להוספת תרופה חדשה.
    ///
    /// החלון מאפשר:
    /// - הזנת שם תרופה
    /// - הזנת מחיר
    /// - הזנת מלאי
    /// </summary>
    private class AddMedicationDialog : Form
    {
        /// <summary>
        /// שדה שם התרופה.
        /// </summary>
        private readonly TextBox _name = new();

        /// <summary>
        /// שדה מחיר התרופה.
        ///
        /// NumericUpDown מאפשר רק מספרים.
        /// </summary>
        private readonly NumericUpDown _price = new();

        /// <summary>
        /// שדה כמות מלאי.
        ///
        /// גם כאן NumericUpDown מונע הכנסת טקסט.
        /// </summary>
        private readonly NumericUpDown _stock = new();

        /// <summary>
        /// התרופה שנוצרה בסיום החלון.
        /// </summary>
        public Medication? Result { get; private set; }

        /// <summary>
        /// פעולה בונה של חלון הוספת תרופה.
        /// </summary>
        public AddMedicationDialog()
        {
            // כותרת החלון
            Text = "Add Medication";

            // גודל החלון
            Size = new Size(380, 260);

            // פתיחה במרכז
            StartPosition = FormStartPosition.CenterParent;

            // צבע רקע
            BackColor = Color.FromArgb(245, 248, 252);

            // פונט כללי
            Font = new Font("Segoe UI", 10F);

            // חלון קבוע
            FormBorderStyle = FormBorderStyle.FixedDialog;

            // ביטול אפשרות הגדלה
            MaximizeBox = false;

            // יצירת שורת שם התרופה
            AddRow("Name:", _name, 20);

            /// <summary>
            /// תווית מחיר.
            /// </summary>
            var lblPrice = new Label
            {
                Text = "Price (₪):",
                Location = new Point(20, 63),
                AutoSize = true
            };

            // מיקום וגודל שדה המחיר
            _price.Location = new Point(150, 60);
            _price.Size = new Size(180, 25);

            // מספר ספרות אחרי הנקודה
            _price.DecimalPlaces = 2;

            // מחיר מקסימלי
            _price.Maximum = 100000;

            // ערך ברירת מחדל
            _price.Value = 50;

            // הוספת שדה המחיר
            Controls.Add(lblPrice);
            Controls.Add(_price);

            /// <summary>
            /// תווית מלאי.
            /// </summary>
            var lblStock = new Label
            {
                Text = "Stock quantity:",
                Location = new Point(20, 103),
                AutoSize = true
            };

            // מיקום וגודל שדה המלאי
            _stock.Location = new Point(150, 100);
            _stock.Size = new Size(180, 25);

            // מלאי מקסימלי
            _stock.Maximum = 100000;

            // מלאי ברירת מחדל
            _stock.Value = 10;

            // הוספת שדה המלאי
            Controls.Add(lblStock);
            Controls.Add(_stock);

            /// <summary>
            /// כפתור שמירת תרופה.
            /// </summary>
            var save = new Button
            {
                Text = "Save",
                Location = new Point(140, 160),
                Size = new Size(90, 32),
                BackColor = Color.FromArgb(40, 160, 90),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
            };

            // הסרת מסגרת
            save.FlatAppearance.BorderSize = 0;

            // בעת לחיצה → שמירת התרופה
            save.Click += OnSave;

            /// <summary>
            /// כפתור ביטול.
            /// </summary>
            var cancel = new Button
            {
                Text = "Cancel",
                Location = new Point(240, 160),
                Size = new Size(90, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(220, 220, 220),
            };

            // הסרת מסגרת
            cancel.FlatAppearance.BorderSize = 0;

            // סגירת החלון ללא שמירה
            cancel.Click += (_, _) => Close();

            // הוספת הכפתורים לחלון
            Controls.Add(save);
            Controls.Add(cancel);
        }

        /// <summary>
        /// פונקציית עזר שיוצרת שורת קלט.
        ///
        /// הפעולה יוצרת:
        /// - Label
        /// - Control מתאים
        /// וממקמת אותם בחלון.
        /// </summary>
        private void AddRow(string text, Control input, int y)
        {
            // יצירת תווית
            var lbl = new Label
            {
                Text = text,
                Location = new Point(20, y + 3),
                AutoSize = true
            };

            // מיקום שדה הקלט
            input.Location = new Point(150, y);

            // גודל שדה הקלט
            input.Size = new Size(180, 25);

            // הוספת הרכיבים לחלון
            Controls.Add(lbl);
            Controls.Add(input);
        }

        /// <summary>
        /// מופעלת כאשר המשתמש לוחץ Save.
        ///
        /// הפעולה:
        /// 1. מבצעת validation על שם התרופה.
        /// 2. יוצרת אובייקט Medication חדש.
        /// 3. סוגרת את החלון בהצלחה.
        /// </summary>
        private void OnSave(object? sender, EventArgs e)
        {
            // validation: שם חובה
            if (string.IsNullOrWhiteSpace(_name.Text))
            {
                MessageBox.Show(
                    "Name is required.",
                    "Validation error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            /// <summary>
            /// יצירת אובייקט התרופה החדש.
            /// </summary>
            Result = new Medication
            {
                // שם התרופה
                Name = _name.Text.Trim(),

                // מחיר התרופה
                Price = _price.Value,

                // כמות מלאי
                StockQuantity = (int)_stock.Value,
            };

            // סימון שהשמירה הצליחה
            DialogResult = DialogResult.OK;

            // סגירת החלון
            Close();
        }
    }

    /// <summary>
    /// חלון עריכת תרופה קיימת.
    /// </summary>
    private class EditMedicationDialog : Form
    {
        /// <summary>
        /// התרופה שנערכת.
        /// </summary>
        private readonly Medication _medication;

        /// <summary>
        /// שדה מחיר.
        /// </summary>
        private readonly NumericUpDown _price = new();

        /// <summary>
        /// שדה מלאי.
        /// </summary>
        private readonly NumericUpDown _stock = new();

        /// <summary>
        /// פעולה בונה של חלון העריכה.
        /// </summary>
        public EditMedicationDialog(Medication medication)
        {
            // שמירת התרופה
            _medication = medication;

            // כותרת החלון
            Text = "Edit Medication";

            // גודל החלון
            Size = new Size(380, 240);

            // פתיחה במרכז
            StartPosition = FormStartPosition.CenterParent;

            // צבע רקע
            BackColor = Color.FromArgb(245, 248, 252);

            // פונט
            Font = new Font("Segoe UI", 10F);

            // חלון קבוע
            FormBorderStyle = FormBorderStyle.FixedDialog;

            // ביטול אפשרות הגדלה
            MaximizeBox = false;

            // שם התרופה
            var lblName = new Label
            {
                Text = $"Medication: {medication.Name}",
                Location = new Point(20, 20),
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };

            Controls.Add(lblName);

            // תווית מחיר
            var lblPrice = new Label
            {
                Text = "Price (₪):",
                Location = new Point(20, 70),
                AutoSize = true
            };

            Controls.Add(lblPrice);

            // שדה מחיר
            _price.Location = new Point(150, 70);
            _price.Size = new Size(180, 25);
            _price.DecimalPlaces = 2;
            _price.Maximum = 100000;
            _price.Value = medication.Price;

            Controls.Add(_price);

            // תווית מלאי
            var lblStock = new Label
            {
                Text = "Stock quantity:",
                Location = new Point(20, 110),
                AutoSize = true
            };

            Controls.Add(lblStock);

            // שדה מלאי
            _stock.Location = new Point(150, 110);
            _stock.Size = new Size(180, 25);
            _stock.Maximum = 100000;
            _stock.Value = medication.StockQuantity;

            Controls.Add(_stock);

            // כפתור שמירה
            var save = new Button
            {
                Text = "Save",
                Location = new Point(140, 160),
                Size = new Size(90, 32),
                BackColor = Color.FromArgb(40, 160, 90),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
            };

            save.FlatAppearance.BorderSize = 0;

            save.Click += OnSave;

            Controls.Add(save);
        }

        /// <summary>
        /// שמירת השינויים בתרופה.
        /// </summary>
        private void OnSave(object? sender, EventArgs e)
        {
            // עדכון מחיר
            _medication.Price = _price.Value;

            // עדכון מלאי
            _medication.StockQuantity = (int)_stock.Value;

            // הצלחה
            DialogResult = DialogResult.OK;

            // סגירת החלון
            Close();
        }
    }
}