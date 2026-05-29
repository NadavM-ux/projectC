using System.Drawing;
using System.IO;
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

    // כפתור פתיחת ביקור חדש
    private readonly Button _addBtn = new();

    // כפתור חזרה לתפריט הראשי
    private readonly Button _backBtn = new();

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

        // גודל התחלתי של החלון
        Size = new Size(960, 600);

        // מאפשר הגדלה והקטנה של החלון
        FormBorderStyle = FormBorderStyle.Sizable;

        // מאפשר maximize
        MaximizeBox = true;

        // גודל מינימלי לחלון
        MinimumSize = new Size(960, 600);

        // פתיחת החלון במרכז ביחס לחלון האב
        StartPosition = FormStartPosition.CenterParent;

        // צבע רקע כללי של המסך
        BackColor = Color.FromArgb(245, 248, 252);

        // הגדרת פונט ברירת מחדל למסך
        Font = new Font("Segoe UI", 10F);


        // Header עליון של המסך
        var header = new Panel
        {
            Dock = DockStyle.Top,
            Height = 95,
            BackColor = Color.White,
        };


        var headerDogPic = new PictureBox
        {
            Location = new Point(25, 12),
            Size = new Size(65, 65),
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.Transparent
        };

        string headerDogPath = Path.Combine(
            Application.StartupPath,
            "Images",
            "dog1.png");

        if (File.Exists(headerDogPath))
        {
            headerDogPic.Image = Image.FromFile(headerDogPath);
        }

        header.Controls.Add(headerDogPic);

        // כותרת ראשית
        var title = new Label
        {
            Text = "Visits & Treatments",
            Font = new Font("Segoe UI", 22F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            AutoSize = true,
            Location = new Point(105, 18),
        };

        // תת כותרת
        var subtitle = new Label
        {
            Text = "Manage clinic visits and treatments",
            Font = new Font("Segoe UI", 11F, FontStyle.Italic),
            ForeColor = Color.FromArgb(100, 110, 130),
            AutoSize = true,
            Location = new Point(108, 58),
        };

        // הוספת רכיבים ל־Header
        header.Controls.Add(title);
        header.Controls.Add(subtitle);

        var headerLine = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 2,
            BackColor = Color.FromArgb(40, 90, 150)
        };

        header.Controls.Add(headerLine);

        // הוספת ה־Header למסך
        Controls.Add(header);

        // לא מציגים התראת חיסונים כללית במסך הראשי,
        // כי ההתראה מוצגת בתוך Open New Visit לפי החיה שנבחרה.
        _alertLabel.Visible = false;

        // מיקום טבלת הביקורים
        _grid.Location = new Point(20, 115);

        // גודל התחלתי
        _grid.Size = new Size(900, 375);

        // הטבלה תגדל עם החלון
        _grid.Anchor = AnchorStyles.Top
                     | AnchorStyles.Bottom
                     | AnchorStyles.Left
                     | AnchorStyles.Right;

        // ביטול יצירת עמודות אוטומטית כדי להגדיר עמודות ידנית
        _grid.AutoGenerateColumns = false;

        // מניעת הוספת שורות ידנית על ידי המשתמש
        _grid.AllowUserToAddRows = false;

        // מניעת שינוי רוחב עמודות על ידי המשתמש
        _grid.AllowUserToResizeColumns = false;

        // מניעת שינוי גובה שורות על ידי המשתמש
        _grid.AllowUserToResizeRows = false;

        // הטבלה היא לקריאה בלבד
        _grid.ReadOnly = true;

        // בחירת שורה מלאה בלחיצה
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

        // צבע רקע של הטבלה
        _grid.BackgroundColor = Color.White;

        // הסתרת עמודות הצד
        _grid.RowHeadersVisible = false;

        // ביטול מסגרת ברירת מחדל
        _grid.BorderStyle = BorderStyle.None;

        // קווי הפרדה עדינים בין שורות
        _grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;

        // שימוש בעיצוב מותאם אישית לכותרות
        _grid.EnableHeadersVisualStyles = false;

        // צבע רקע של כותרות העמודות
        _grid.ColumnHeadersDefaultCellStyle.BackColor =
            Color.FromArgb(40, 90, 150);

        // צבע טקסט של הכותרות
        _grid.ColumnHeadersDefaultCellStyle.ForeColor =
            Color.White;

        // פונט כותרות
        _grid.ColumnHeadersDefaultCellStyle.Font =
            new Font("Segoe UI", 11F, FontStyle.Bold);

        // גובה כותרות
        _grid.ColumnHeadersHeight = 42;

        // גובה שורות
        _grid.RowTemplate.Height = 34;

        // צבע שורות מתחלף
        _grid.AlternatingRowsDefaultCellStyle.BackColor =
            Color.FromArgb(245, 248, 252);

        // בחירת שורה מלאה
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

        // ביטול בחירת כמה שורות יחד
        _grid.MultiSelect = false;

        // צבע בחירת שורה
        _grid.DefaultCellStyle.SelectionBackColor =
            Color.FromArgb(90, 140, 220);

        // צבע טקסט של שורה נבחרת
        _grid.DefaultCellStyle.SelectionForeColor =
            Color.White;

        // התאמה אוטומטית של העמודות לרוחב הטבלה
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        // עמודת תאריך הביקור
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Date",
            DataPropertyName = nameof(Visit.VisitDateTime),
            DefaultCellStyle = { Format = "dd-MM-yyyy HH:mm" }
        });

        // עמודת מספר השבב של החיה
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Animal",
            DataPropertyName = nameof(Visit.AnimalChipNumber)
        });

        // עמודת שם המשתמש של הווטרינר
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Vet",
            DataPropertyName = nameof(Visit.VetUsername)
        });

        // עמודת סיבת ההגעה
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Reason",
            DataPropertyName = nameof(Visit.Reason)
        });

        // עמודת האבחנה
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Diagnosis",
            DataPropertyName = nameof(Visit.Diagnosis)
        });

        // עמודת המחיר הכולל של הביקור
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Total (₪)",
            DataPropertyName = nameof(Visit.TotalPrice)
        });

        // אירוע לחיצה כפולה על ביקור
        _grid.CellDoubleClick += OnVisitDoubleClick;

        // הוספת הטבלה למסך
        Controls.Add(_grid);

        // כפתור לפתיחת ביקור חדש
        _addBtn.Text = "➕ Open New Visit";
        _addBtn.Location = new Point(20, 495);
        _addBtn.Size = new Size(220, 48);
        _addBtn.BackColor = Color.FromArgb(40, 160, 90);
        _addBtn.ForeColor = Color.White;
        _addBtn.FlatStyle = FlatStyle.Flat;
        _addBtn.Font = new Font("Segoe UI", 11F, FontStyle.Bold);

        _addBtn.FlatAppearance.BorderSize = 0;
        _addBtn.Click += OnNewVisit;
        _addBtn.Anchor = AnchorStyles.Bottom;
        Controls.Add(_addBtn);

        _backBtn.Text = "← Main Menu";
        _backBtn.Location = new Point(260, 495);
        _backBtn.Size = new Size(220, 48);
        _backBtn.BackColor = Color.FromArgb(120, 130, 150);
        _backBtn.ForeColor = Color.White;
        _backBtn.FlatStyle = FlatStyle.Flat;
        _backBtn.Font = new Font("Segoe UI", 11F, FontStyle.Bold);

        _backBtn.FlatAppearance.BorderSize = 0;
        _backBtn.Click += (_, _) => Close();
        _backBtn.Anchor = AnchorStyles.Bottom;
        Controls.Add(_backBtn);

        _addBtn.Location =
           new Point((ClientSize.Width / 2) - 320,
                     ClientSize.Height - 55);

        _backBtn.Location =
            new Point((ClientSize.Width / 2) + 20,
                      ClientSize.Height - 55);

        // אירוע שמופעל כאשר גודל החלון משתנה
        Resize += VisitsForm_Resize;


    }

    // פונקציה שמתאימה את גודל הטקסט והשורות לפי גודל החלון
    private void VisitsForm_Resize(object? sender, EventArgs e)
    {
        if (Width > 1400)
        {
            // הגדלת פונט הטבלה
            _grid.DefaultCellStyle.Font =
                new Font("Segoe UI", 14F);

            // הגדלת פונט הכותרות
            _grid.ColumnHeadersDefaultCellStyle.Font =
                new Font("Segoe UI", 15F, FontStyle.Bold);

            // הגדלת גובה כותרות
            _grid.ColumnHeadersHeight = 60;

            // הגדלת גובה שורות
            foreach (DataGridViewRow row in _grid.Rows)
            {
                row.Height = 55;
            }

            // הגדלת כפתורים
            _addBtn.Size = new Size(300, 70);
            _backBtn.Size = new Size(300, 70);

            _addBtn.Font =
                new Font("Segoe UI", 14F, FontStyle.Bold);

            _backBtn.Font =
                new Font("Segoe UI", 14F, FontStyle.Bold);
        }
        else
        {
            // גודל רגיל לטבלה
            _grid.DefaultCellStyle.Font =
                new Font("Segoe UI", 10F);

            // גודל רגיל לכותרות
            _grid.ColumnHeadersDefaultCellStyle.Font =
                new Font("Segoe UI", 11F, FontStyle.Bold);

            // גובה רגיל לכותרות
            _grid.ColumnHeadersHeight = 42;

            // גובה רגיל לשורות
            foreach (DataGridViewRow row in _grid.Rows)
            {
                row.Height = 34;
            }

            // גודל רגיל לכפתורים
            _addBtn.Size = new Size(220, 48);
            _backBtn.Size = new Size(220, 48);

            _addBtn.Font =
                new Font("Segoe UI", 11F, FontStyle.Bold);

            _backBtn.Font =
                new Font("Segoe UI", 11F, FontStyle.Bold);
        }

        // מיקום הכפתורים במרכז
        _addBtn.Location =
            new Point((ClientSize.Width / 2) - _addBtn.Width - 20,
                      ClientSize.Height - 55);

        _backBtn.Location =
            new Point((ClientSize.Width / 2) + 20,
                      ClientSize.Height - 55);
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


    // פתיחת חלון פרטי ביקור בלחיצה כפולה
    private void OnVisitDoubleClick(object? sender, DataGridViewCellEventArgs e)
    {
        // אם לא נבחרה שורה תקינה
        if (e.RowIndex < 0)
            return;

        // שליפת הביקור שנבחר
        var visit = _grid.Rows[e.RowIndex].DataBoundItem as Visit;

        // אם לא נמצא ביקור
        if (visit == null)
            return;

        // יצירת חלון פרטי ביקור
        var details = new Form
        {
            Text = "Visit Details",
            Size = new Size(500, 500),
            StartPosition = FormStartPosition.CenterParent,
            BackColor = Color.White,
            Font = new Font("Segoe UI", 10F),
            FormBorderStyle = FormBorderStyle.Sizable,
            MinimumSize = new Size(500, 500),
            AutoScroll = true,
        };




        // כרטיס פרטי ביקור - ישמש רק כשמגדילים את החלון
        var detailsCard = new Panel
        {
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Location = new Point(20, 20),
            Size = new Size(440, 340)
        };

        // כותרת חיצונית למסך גדול בלבד
        var pageTitle = new Label
        {
            Text = "Visit Details",
            Font = new Font("Segoe UI", 28F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            AutoSize = true,
            Visible = false
        };

        var pageSubtitle = new Label
        {
            Text = "View full information about the selected medical visit",
            Font = new Font("Segoe UI", 12F, FontStyle.Italic),
            ForeColor = Color.FromArgb(100, 110, 130),
            AutoSize = true,
            Visible = false
        };


        var detailsTitle = new Label
        {
            Text = "🩺 Visit Information",
            Font = new Font("Segoe UI", 20F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            AutoSize = true,
            Location = new Point(30, 25)
        };


        var detailsSubtitle = new Label
        {
            Text = "Full medical visit information",
            Font = new Font("Segoe UI", 10F, FontStyle.Italic),
            ForeColor = Color.FromArgb(100, 110, 130),
            AutoSize = true,
            Location = new Point(33, 65)
        };

        // אזור טקסט של פרטי הביקור
        var info = new RichTextBox
        {
            Multiline = true,
            ReadOnly = true,
            DetectUrls = false,
            ShortcutsEnabled = false,
            ScrollBars = RichTextBoxScrollBars.Vertical,
            BorderStyle = BorderStyle.None,
            TabStop = false,
            Cursor = Cursors.Default,
            BackColor = Color.White,
            Font = new Font("Segoe UI", 11F),
            Location = new Point(20, 20),
            Size = new Size(440, 340),

            Anchor =
                AnchorStyles.Top |
                AnchorStyles.Bottom |
                AnchorStyles.Left |
                AnchorStyles.Right,

            Text =
                $"🩺 Visit Details\r\n\r\n" +
                $"📅 Date: {visit.VisitDateTime:dd-MM-yyyy HH:mm}\r\n\r\n" +
                $"🐾 Animal Chip: {visit.AnimalChipNumber}\r\n\r\n" +
                $"👨‍⚕️ Vet: {visit.VetUsername}\r\n\r\n" +
                $"📋 Reason:\r\n{visit.Reason}\r\n\r\n" +
                $"🔍 Diagnosis:\r\n{visit.Diagnosis}\r\n\r\n" +
                $"💊 Medications:\r\n" +
                $"{string.Join(", ", visit.MedicationsGiven)}\r\n\r\n" +
                $"💰 Total Price: ₪{visit.TotalPrice:0.##}"
        };



        // כפתור סגירה
        var closeBtn = new Button
        {
            Text = "Close",
            Size = new Size(140, 45),
            Location = new Point(180, 390),
            Font = new Font("Segoe UI", 11F, FontStyle.Bold),
            Anchor = AnchorStyles.Bottom,
            BackColor = Color.FromArgb(40, 90, 150),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };

        closeBtn.FlatAppearance.BorderSize = 0;

        closeBtn.Click += (_, _) => details.Close();

        // הוספת רכיבים לחלון
        detailsCard.Controls.Add(detailsTitle);
        detailsCard.Controls.Add(detailsSubtitle);
        detailsCard.Controls.Add(info);
        detailsCard.Controls.Add(closeBtn);

        details.Controls.Add(pageTitle);
        details.Controls.Add(pageSubtitle);
        details.Controls.Add(detailsCard);

        info.TabStop = false;


        // מניעת focus מה־RichTextBox
        details.Shown += (_, _) =>
        {
            closeBtn.Focus();
        };

        // אם ה־RichTextBox מקבל focus בטעות
        info.Enter += (_, _) =>
        {
            closeBtn.Focus();
        };

        // ביטול פוקוס מהטקסט כדי שלא יופיע סמן כתיבה
        details.Shown += (_, _) =>
        {
            closeBtn.Focus();
        };

        // פונקציה שמסדרת את חלון הפרטים לפי גודל החלון
        void ApplyDetailsLayout()
        {
            // מצב גדול בלבד
            if (details.ClientSize.Width > 900)
            {
                details.BackColor = Color.FromArgb(225, 235, 248);

                pageTitle.Visible = true;
                pageSubtitle.Visible = true;

                detailsTitle.Visible = true;
                detailsSubtitle.Visible = true;

                detailsCard.BorderStyle = BorderStyle.None;
                detailsCard.BackColor = Color.White;

                detailsCard.Size = new Size(
                    Math.Min(1500, Math.Max(1100, (int)(details.ClientSize.Width * 0.78))),
                    Math.Min(850, Math.Max(650, (int)(details.ClientSize.Height * 0.72)))
                );

                detailsCard.Location = new Point(
                    (details.ClientSize.Width - detailsCard.Width) / 2,
                    (details.ClientSize.Height - detailsCard.Height) / 2
                );
                pageTitle.Location = new Point(
                       (details.ClientSize.Width - pageTitle.Width) / 2,
                            detailsCard.Top - 95
                );

                pageSubtitle.Location = new Point(
                    (details.ClientSize.Width - pageSubtitle.Width) / 2,
                    detailsCard.Top - 45
                );
                detailsTitle.Location = new Point(45, 30);
                detailsSubtitle.Location = new Point(48, 78);

                info.Location = new Point(45, 125);
                info.Size = new Size(
                    detailsCard.Width - 90,
                    detailsCard.Height - 220
                );

                info.Font = new Font("Segoe UI", 16F);
                info.BackColor = Color.White;

                closeBtn.Size = new Size(170, 55);
                closeBtn.Font = new Font("Segoe UI", 13F, FontStyle.Bold);

                closeBtn.Location = new Point(
                    (detailsCard.Width - closeBtn.Width) / 2,
                    detailsCard.Height - 85
                );
            }
            else
            {
                // מצב קטן - מחזירים בדיוק לעיצוב המקורי
                details.BackColor = Color.White;

                pageTitle.Visible = false;
                pageSubtitle.Visible = false;

                detailsTitle.Visible = false;
                detailsSubtitle.Visible = false;

                detailsCard.BorderStyle = BorderStyle.None;
                detailsCard.Location = new Point(0, 0);
                detailsCard.Size = details.ClientSize;
                detailsCard.BackColor = Color.White;

                info.Location = new Point(20, 20);
                info.Size = new Size(
                    Math.Max(100, details.ClientSize.Width - 40),
                    Math.Max(100, details.ClientSize.Height - 120)
                );

                info.Font = new Font("Segoe UI", 11F);
                info.BackColor = Color.White;

                closeBtn.Size = new Size(140, 45);
                closeBtn.Font = new Font("Segoe UI", 11F, FontStyle.Bold);

                closeBtn.Location = new Point(
                    (details.ClientSize.Width - closeBtn.Width) / 2,
                    details.ClientSize.Height - 70
                );
            }
        }

        // בכל שינוי גודל נפעיל את סידור החלון
        details.Resize += (_, _) =>
        {
            ApplyDetailsLayout();
        };

        // הפעלה ראשונית כדי שגם בפתיחה הכול יהיה במקום
        ApplyDetailsLayout();

        details.PerformLayout();
        PerformLayout();

        // פתיחת החלון
        details.ShowDialog(this);
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


            // מעבר על כל התרופות שניתנו במהלך הביקור והורדת הכמות מהמלאי
            foreach (var medicationQuantity in dlg.MedicationQuantities)
            {
                // שם התרופה
                var medName = medicationQuantity.Key;

                // הכמות שניתנה
                var quantity = medicationQuantity.Value;

                // חיפוש התרופה במערכת לפי שם
                var medication = _store.Medications
                    .FirstOrDefault(m => m.Name == medName);

                // אם התרופה נמצאה
                if (medication != null)
                {
                    // הורדת הכמות שניתנה מהמלאי
                    medication.StockQuantity -= quantity;

                    // הגנה כדי שהמלאי לא יהיה שלילי
                    if (medication.StockQuantity < 0)
                    {
                        medication.StockQuantity = 0;
                    }
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
        private readonly RichTextBox _reason = new();


        /// <summary>
        /// שדה טקסט לאבחנה רפואית.
        /// </summary>
        private readonly RichTextBox _diagnosis = new();



        // שומר האם כרגע הטקסט של Reason מיושר כעברית
        private bool _reasonHebrewAlignment = false;

        // שומר האם כרגע הטקסט של Diagnosis מיושר כעברית
        private bool _diagnosisHebrewAlignment = false;




        /// <summary>
        /// רשימת תרופות עם אפשרות סימון.
        /// כל תרופה מוצגת עם השם והמחיר שלה.
        /// </summary>
        private readonly CheckedListBox _meds = new();

        // תיבת חיפוש תרופה
        private readonly TextBox _medSearch = new();

        // רשימת התרופות שנבחרו + כמות לכל תרופה
        // לדוגמה: Antibiotics -> 3
        private readonly Dictionary<string, int> _selectedMedicationQuantities = new();

        // משתנה עזר כדי לדעת מתי הרשימה נטענת מחדש
        private bool _loadingMedicationList = false;

        /// <summary>
        /// תווית המציגה את המחיר הכולל של הביקור.
        /// </summary>
        private readonly Label _totalLabel = new();

        /// <summary>
        /// סימון האם הביקור הנוכחי הוא חיסון של היום.
        /// אם מסומן, תאריך החיסון האחרון של החיה יעודכן.
        /// </summary>
        private readonly CheckBox _markVaccinated = new();


        // תווית התראה על צורך בחיסון
        private readonly Label _vaccinationAlert = new();

        /// <summary>
        /// תוצאת החלון: הביקור החדש שנוצר.
        /// אם המשתמש ביטל את הפעולה, הערך יישאר null.
        /// </summary>
        public Visit? Result { get; private set; }

        /// <summary>
        /// כמויות התרופות שנבחרו בביקור.
        /// המפתח הוא שם התרופה, והערך הוא הכמות שניתנה.
        /// </summary>
        public Dictionary<string, int> MedicationQuantities { get; private set; } = new();

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
            Size = new Size(560, 780);

            // פתיחת החלון במרכז
            StartPosition = FormStartPosition.CenterParent;

            // צבע רקע
            BackColor = Color.FromArgb(245, 248, 252);

            // פונט כללי
            Font = new Font("Segoe UI", 10F);

            // רקע כחול שיופיע רק במסך הגדול
            var outerBackground = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(225, 235, 248),
                Visible = false
            };
            Controls.Add(outerBackground);
            outerBackground.SendToBack();

            // כותרת מעוצבת למסך גדול בלבד
            var bigHeader = new Panel
            {
                Size = new Size(950, 110),
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                Visible = false
            };

            var bigTitle = new Label
            {
                Text = "🐾 New Visit & Treatment",
                Font = new Font("Segoe UI", 26F, FontStyle.Bold),
                ForeColor = Color.FromArgb(40, 90, 150),
                AutoSize = true,
                Location = new Point(145, 22)
            };

            var bigSubtitle = new Label
            {
                Text = "Create a new visit, diagnosis and medications for your patient",
                Font = new Font("Segoe UI", 12F, FontStyle.Italic),
                ForeColor = Color.FromArgb(100, 110, 130),
                AutoSize = true,
                Location = new Point(148, 75)
            };

            var leftAnimalPic = CreateDecorativeAnimalPicture(
                "dog1.png",
                new Point(20, 15),
                new Size(75, 75));



            bigHeader.Controls.Add(bigTitle);
            bigHeader.Controls.Add(bigSubtitle);
            bigHeader.Controls.Add(leftAnimalPic);


            Controls.Add(bigHeader);
            bigHeader.BringToFront();

            // פאנל מרכזי שמחזיק את כל הטופס
            var contentPanel = new Panel
            {
                Size = new Size(560, 720),
                BackColor = Color.Transparent
            };

            // מיקום התחלתי במרכז החלון
            contentPanel.Location = new Point(
                (ClientSize.Width - contentPanel.Width) / 2,
                bigHeader.Bottom + 25
            );

            // הפאנל נשאר ממורכז
            contentPanel.Anchor = AnchorStyles.Top;

            // חלון דיאלוג קבוע
            FormBorderStyle = FormBorderStyle.Sizable;

            // ביטול כפתור הגדלה
            MaximizeBox = true;

            // תווית חיפוש צ'יפ
            var lblChip = new Label
            {
                Text = "Chip number:",
                Location = new Point(20, 23),
                AutoSize = true
            };

            // כפתור חיפוש חיה לפי צ'יפ
            var searchChipBtn = new Button
            {
                Text = "🔍 Search animal by chip",
                Location = new Point(20, 20),
                Size = new Size(220, 36),
                BackColor = Color.FromArgb(40, 90, 150),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };

            searchChipBtn.FlatAppearance.BorderSize = 0;
            searchChipBtn.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Left;

            searchChipBtn.Click += OnSearchChip;

            contentPanel.Controls.Add(searchChipBtn);


            // תווית בחירת חיה
            var lblAnimal = new Label { Text = "Animal:", Location = new Point(20, 75), AutoSize = true };


            // מיקום וגודל רשימת החיות
            _animal.Location = new Point(150, 72);
            _animal.Size = new Size(370, 25);
            _animal.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Left |
                AnchorStyles.Right;

            _animal.MaximumSize = new Size(900, 25);

            // המשתמש יכול לבחור רק מתוך הרשימה ולא להקליד חיה ידנית
            _animal.DropDownStyle = ComboBoxStyle.DropDownList;

            // טעינת כל החיות מתוך המאגר אל הרשימה
            foreach (var a in _store.Animals)
            {
                _animal.Items.Add($"{a.Name} (chip {a.ChipNumber}) — {a.Species}");
            }

            // לא בוחרים חיה אוטומטית.
            // הווטרינר חייב לבחור חיה בעצמו מתוך הרשימה.
            _animal.SelectedIndex = -1;

            // הוספת התווית ורשימת החיות לחלון
            contentPanel.Controls.Add(lblAnimal);
            contentPanel.Controls.Add(_animal);

            // מיקום תווית ההתראה
            _vaccinationAlert.Location = new Point(20, 605);

            // צבע אדום להתראה
            _vaccinationAlert.ForeColor = Color.DarkRed;

            // פונט מודגש
            _vaccinationAlert.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

            // גודל אוטומטי
            _vaccinationAlert.AutoSize = true;
            _vaccinationAlert.Anchor =
                AnchorStyles.Left |
                AnchorStyles.Bottom;

            // הוספת ההתראה למסך
            contentPanel.Controls.Add(_vaccinationAlert);

            // בדיקה ראשונית של החיה שנבחרה
            UpdateVaccinationAlert();

            // בכל החלפת חיה נבדוק מחדש
            _animal.SelectedIndexChanged += (_, _) => UpdateVaccinationAlert();



            // תווית תאריך ושעה
            var lblDate = new Label
            {
                Text = "Date / time:",
                Location = new Point(20, 103),
                AutoSize = true
            };

            // מיקום וגודל בוחר התאריך
            _date.Location = new Point(150, 100);
            _date.Size = new Size(370, 25);
            _date.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Left |
                AnchorStyles.Right;

            _date.MaximumSize = new Size(900, 25);

            // שימוש בפורמט מותאם אישית
            _date.Format = DateTimePickerFormat.Custom;

            // פורמט תאריך:
            // יום - חודש - שנה
            // שעה:דקות
            _date.CustomFormat = "dd-MM-yyyy HH:mm";

            // ערך ברירת מחדל: התאריך והשעה הנוכחיים
            _date.Value = DateTime.Now;

            // המשתמש לא יכול לשנות את התאריך
            _date.Enabled = false;

            // הוספת רכיבי התאריך לחלון
            contentPanel.Controls.Add(lblDate);
            contentPanel.Controls.Add(_date);

            // תווית סיבת ביקור
            var lblReason = new Label { Text = "Reason for visit:", Location = new Point(20, 143), AutoSize = true };

            // מיקום וגודל שדה סיבת הביקור
            _reason.Location = new Point(150, 140);
            _reason.Size = new Size(370, 60);
            _reason.MaximumSize = new Size(900, 60);
            _reason.ScrollBars = RichTextBoxScrollBars.Vertical;
            _reason.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Left |
                AnchorStyles.Right;
            _reason.KeyPress += (_, e) =>
            {
                UpdateTextAlignmentByKey(_reason, e.KeyChar, ref _reasonHebrewAlignment);
            };



            // הוספת סיבת הביקור לחלון
            contentPanel.Controls.Add(lblReason);
            contentPanel.Controls.Add(_reason);

            // תווית אבחנה
            var lblDiag = new Label { Text = "Diagnosis:", Location = new Point(20, 220), AutoSize = true };

            // מיקום וגודל שדה האבחנה
            _diagnosis.Location = new Point(150, 220);
            _diagnosis.Size = new Size(370, 60);
            _diagnosis.MaximumSize = new Size(900, 60);

            // מאפשר כתיבה בכמה שורות
            _diagnosis.Multiline = true;
            _diagnosis.ScrollBars = RichTextBoxScrollBars.Vertical;
            _diagnosis.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Left |
                AnchorStyles.Right;
            _diagnosis.KeyPress += (_, e) =>
            {
                UpdateTextAlignmentByKey(_diagnosis, e.KeyChar, ref _diagnosisHebrewAlignment);
            };


            // הוספת שדה האבחנה לחלון
            contentPanel.Controls.Add(lblDiag);
            contentPanel.Controls.Add(_diagnosis);
            AddPlaceholder(_reason, "Enter reason here...");
            AddPlaceholder(_diagnosis, "Enter diagnosis here...");

            // כותרת אזור התרופות
            var lblMeds = new Label { Text = "Medications administered:", Location = new Point(20, 320), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold) };
            
            // תיבת חיפוש תרופה
            _medSearch.Location = new Point(260, 318);
            _medSearch.Size = new Size(250, 28);
            _medSearch.Font = new Font("Segoe UI", 10F);
            _medSearch.PlaceholderText = "Search medication...";
            _medSearch.BorderStyle = BorderStyle.FixedSingle;

            // בכל הקלדה נטען מחדש את רשימת התרופות לפי החיפוש
            _medSearch.TextChanged += (_, _) => LoadMedicationList(_medSearch.Text);

            // הוספת תיבת החיפוש לחלון
            contentPanel.Controls.Add(_medSearch);

            // מיקום וגודל רשימת התרופות
            _meds.Location = new Point(20, 350);
            _meds.Size = new Size(500, 180);
            _meds.MaximumSize = new Size(900, 180);
            _meds.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Left |
                AnchorStyles.Right;

            // מאפשר סימון תרופה בלחיצה אחת
            _meds.CheckOnClick = true;

            // טעינת רשימת התרופות בפעם הראשונה
            LoadMedicationList();

            _meds.ItemCheck += (sender, e) =>
            {
                // אם הרשימה נטענת מחדש בעקבות חיפוש, לא מפעילים לוגיקה
                if (_loadingMedicationList)
                    return;

                // הטקסט של התרופה שנבחרה
                var itemText = _meds.Items[e.Index]?.ToString() ?? "";

                // אם התרופה אזלה מהמלאי
                if (itemText.Contains("OUT OF STOCK"))
                {
                    e.NewValue = CheckState.Unchecked;

                    MessageBox.Show(
                        "This medication is out of stock.",
                        "Out of stock",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // חילוץ שם התרופה מתוך הטקסט
                var dashIndex = itemText.IndexOf(" — ", StringComparison.Ordinal);
                var medicationName = dashIndex > 0 ? itemText[..dashIndex] : itemText;

                // אם המשתמש מסמן תרופה
                if (e.NewValue == CheckState.Checked)
                {
                    // חיפוש התרופה במערכת
                    var medication = _store.Medications
                        .FirstOrDefault(m => m.Name == medicationName);

                    if (medication == null)
                    {
                        e.NewValue = CheckState.Unchecked;
                        return;
                    }

                    // פתיחת חלון בחירת כמות
                    int? quantity = AskMedicationQuantity(
                        medication.Name,
                        medication.StockQuantity);

                    // אם המשתמש ביטל, לא מסמנים את התרופה
                    if (quantity == null)
                    {
                        e.NewValue = CheckState.Unchecked;
                        return;
                    }

                    // שמירת הכמות שנבחרה
                    _selectedMedicationQuantities[medication.Name] = quantity.Value;
                }
                else
                {
                    // אם המשתמש הסיר סימון, מוחקים את התרופה מהרשימה
                    _selectedMedicationQuantities.Remove(medicationName);
                }

                // אחרי שהסימון הסתיים, מרעננים את הרשימה ואת המחיר
                BeginInvoke(new Action(() =>
                {
                    LoadMedicationList(_medSearch.Text);
                    UpdateTotal();
                }));
            };

            // הוספת התרופות לחלון
            contentPanel.Controls.Add(lblMeds);
            contentPanel.Controls.Add(_meds); ;

            // טקסט של סימון חיסון
            _markVaccinated.Text = "Mark this visit as today's vaccination";

            // מיקום סימון החיסון
            _markVaccinated.Location = new Point(20, 470);

            // התאמת גודל אוטומטית לטקסט
            _markVaccinated.AutoSize = true;

            // הוספת סימון החיסון לחלון
            contentPanel.Controls.Add(_markVaccinated);

            // הצגת מחיר התחלתי: מחיר בסיס בלבד
            _totalLabel.Text = $"Total: ₪{_basePrice:0.##} (base)";

            // מיקום תווית המחיר
            _totalLabel.Location = new Point(20, 570);

            // התאמת גודל אוטומטית
            _totalLabel.AutoSize = true;
            _totalLabel.Anchor =
                AnchorStyles.Left |
                AnchorStyles.Bottom;

            // פונט מודגש למחיר
            _totalLabel.Font = new Font("Segoe UI", 12F, FontStyle.Bold);

            // צבע המחיר
            _totalLabel.ForeColor = Color.FromArgb(40, 90, 150);

            // הוספת המחיר לחלון
            contentPanel.Controls.Add(_totalLabel);

            // כפתור שמירת ביקור
            var save = new Button
            {
                Text = "Save Visit",
                Location = new Point(350, 645),
                Size = new Size(110, 36),
                BackColor = Color.FromArgb(40, 160, 90),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            };

            // הסרת מסגרת מכפתור השמירה
            save.FlatAppearance.BorderSize = 0;
            save.Anchor =
                AnchorStyles.Bottom;

            // קישור לחיצה לפעולת שמירה
            save.Click += OnSave;

            // כפתור ביטול
            var cancel = new Button
            {
                Text = "Cancel",
                Location = new Point(480, 645),
                Size = new Size(110, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(220, 220, 220),
            };

            // הסרת מסגרת מכפתור הביטול
            cancel.FlatAppearance.BorderSize = 0;
            cancel.Anchor =
                AnchorStyles.Bottom;
            // סגירת החלון ללא שמירה
            cancel.Click += (_, _) => Close();

            // הוספת כפתורי השמירה והביטול
            contentPanel.Controls.Add(save);
            contentPanel.Controls.Add(cancel);

            // הוספת הפאנל לחלון
            Controls.Add(contentPanel);

            // שמירה על מיקום ממורכז בזמן Resize
            Resize += (_, _) =>
            {
                // מצב חלון גדול
                if (ClientSize.Width > 1200)
                {
                    // מצב חלון גדול בלבד

 
                    // הצגת כותרת עליונה מעוצבת
                    bigHeader.Visible = true;

                    bigHeader.Size = new Size(
                        Math.Min(1200, ClientSize.Width - 220),
                        120);

                    bigHeader.Location = new Point(
                        (ClientSize.Width - bigHeader.Width) / 2,
                        25
                    );

                    leftAnimalPic.Location = new Point(25, 12);
                    leftAnimalPic.Size = new Size(95, 95);

                    // מרכזים את התוכן שבתוך הכותרת העליונה
                    int headerContentWidth = 720;

                    int headerContentStartX =
                        (bigHeader.Width - headerContentWidth) / 2;

                    leftAnimalPic.Location = new Point(headerContentStartX, 12);
                    leftAnimalPic.Size = new Size(95, 95);

                    bigTitle.Location = new Point(headerContentStartX + 130, 22);
                    bigSubtitle.Location = new Point(headerContentStartX + 133, 75);



                    // גודל דינמי לפי גודל המסך
                    int panelWidth = (int)(ClientSize.Width * 0.78);
                    int panelHeight = (int)(ClientSize.Height * 0.78);

                    // גבולות כדי שלא יהיה קטן מדי או ענק מדי
                    panelWidth = Math.Max(panelWidth, 1150);
                    panelHeight = Math.Max(panelHeight, 800);

                    contentPanel.Size = new Size(panelWidth, panelHeight);
                    contentPanel.BackColor = Color.White;
                    contentPanel.BorderStyle = BorderStyle.None;

                    // רקע עדין מחוץ למסגרת
                    BackColor = Color.FromArgb(225, 235, 248);
                    outerBackground.BackColor = Color.FromArgb(225, 235, 248);
                    outerBackground.Visible = true;
                    outerBackground.SendToBack();
                    contentPanel.BringToFront();
                    bigHeader.BringToFront();

                    Font = new Font("Segoe UI", 14F);

                    // משתנים לעיצוב מסודר
                    int margin = 60;
                    int labelX = margin;
                    int fieldX = 250;
                    int fieldWidth = contentPanel.Width - fieldX - margin;

                    // Search button
                    searchChipBtn.Location = new Point(margin, 35);
                    searchChipBtn.Size = new Size(330, 60);
                    searchChipBtn.Font =
                        new Font("Segoe UI", 14F, FontStyle.Bold);

                    // Animal
                    lblAnimal.Location = new Point(labelX, 125);
                    lblAnimal.Font = new Font("Segoe UI", 14F);

                    _animal.Location = new Point(fieldX, 118);
                    _animal.Size = new Size(430, 40);
                    _animal.Font = new Font("Segoe UI", 14F);

                    // Date
                    lblDate.Location = new Point(fieldX + 520, 125);
                    lblDate.Font = new Font("Segoe UI", 14F);

                    _date.Location = new Point(fieldX + 660, 123);
                    _date.Size = new Size(360, 40);
                    _date.Font = new Font("Segoe UI", 14F);

                    // Reason — תיקון הדריסה
                    lblReason.Location = new Point(labelX, 215);
                    lblReason.Font = new Font("Segoe UI", 14F);

                    _reason.Location = new Point(fieldX, 205);
                    _reason.MaximumSize = Size.Empty;
                    _reason.Size = new Size(fieldWidth, 170);
                    _reason.Font = new Font("Segoe UI", 14F);
                    _reason.ScrollBars = RichTextBoxScrollBars.Vertical;

                    // Diagnosis
                    lblDiag.Location = new Point(labelX, 420);
                    lblDiag.Font = new Font("Segoe UI", 14F);

                    _diagnosis.Location = new Point(fieldX, 410);
                    _diagnosis.MaximumSize = Size.Empty;
                    _diagnosis.Size = new Size(fieldWidth, 170);
                    _diagnosis.Font = new Font("Segoe UI", 14F);
                    _diagnosis.ScrollBars = RichTextBoxScrollBars.Vertical;

                    // Vaccination checkbox
                    _markVaccinated.Location = new Point(fieldX, 610);
                    _markVaccinated.Font = new Font("Segoe UI", 13F);

                    // Medications title
                    lblMeds.Location = new Point(labelX, 660);
                    lblMeds.Font =
                        new Font("Segoe UI", 14F, FontStyle.Bold);

                    _medSearch.Location = new Point(labelX + 330, 655);
                    _medSearch.Size = new Size(420, 38);
                    _medSearch.Font = new Font("Segoe UI", 13F);

                    // Medications list
                    _meds.Location = new Point(labelX, 705);
                    _meds.MaximumSize = Size.Empty;

                    // גובה דינמי: התרופות יורדות עד קצת מעל אזור המחיר והכפתורים
                    int medsBottom = contentPanel.Height - 125;
                    int medsHeight = Math.Max(330, medsBottom - _meds.Location.Y);

                    _meds.Size = new Size(contentPanel.Width - (margin * 2), medsHeight);
                    _meds.Font = new Font("Segoe UI", 14F);

                    // Total
                    _totalLabel.Location = new Point(labelX, contentPanel.Height - 85);
                    _totalLabel.Font =
                        new Font("Segoe UI", 18F, FontStyle.Bold);

                    // Alert - במסך גדול ההערה תופיע מעל המחיר כדי שלא תדרוס אותו
                    _vaccinationAlert.Location = new Point(labelX, contentPanel.Height - 125);
                    _vaccinationAlert.Font =
                        new Font("Segoe UI", 13F, FontStyle.Bold);
                    _vaccinationAlert.BringToFront();

                    // Buttons
                    save.Size = new Size(190, 62);
                    cancel.Size = new Size(190, 62);

                    save.Font =
                        new Font("Segoe UI", 14F, FontStyle.Bold);

                    cancel.Font =
                        new Font("Segoe UI", 14F);

                    // מיקום כפתורים בתחתית ימין של הכרטיס
                    save.Location = new Point(contentPanel.Width - 460, contentPanel.Height - 95);
                    cancel.Location = new Point(contentPanel.Width - 240, contentPanel.Height - 95);
                }

                // מצב חלון רגיל
                else
                {
                    // החזרת המסך הקטן למצב המקורי
                    bigHeader.Visible = false;

                    BackColor = Color.FromArgb(245, 248, 252);
                    outerBackground.Visible = false;

                    Font = new Font("Segoe UI", 10F);

                    // חשוב: קודם מחזירים את גודל הפאנל
                    contentPanel.Size = new Size(560, 720);
                    contentPanel.Location = new Point(0, 0);
                    contentPanel.BackColor = Color.Transparent;
                    contentPanel.BorderStyle = BorderStyle.None;

                    // כפתור חיפוש - מצב חלון קטן
                    searchChipBtn.Location = new Point(20, 20);
                    searchChipBtn.Size = new Size(220, 36);
                    searchChipBtn.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
                    searchChipBtn.TextAlign = ContentAlignment.MiddleCenter;

                    // Animal
                    lblAnimal.Location = new Point(20, 75);
                    lblAnimal.Font = new Font("Segoe UI", 10F);

                    _animal.Location = new Point(150, 72);
                    _animal.Size = new Size(370, 25);
                    _animal.Font = new Font("Segoe UI", 10F);

                    // Date
                    lblDate.Location = new Point(20, 103);
                    lblDate.Font = new Font("Segoe UI", 10F);

                    _date.Location = new Point(150, 100);
                    _date.Size = new Size(370, 25);
                    _date.Font = new Font("Segoe UI", 10F);

                    // Reason
                    lblReason.Location = new Point(20, 143);
                    lblReason.Font = new Font("Segoe UI", 10F);

                    _reason.Location = new Point(150, 140);
                    _reason.MaximumSize = new Size(900, 60);
                    _reason.Size = new Size(370, 60);
                    _reason.Font = new Font("Segoe UI", 10F);

                    // Diagnosis
                    lblDiag.Location = new Point(20, 220);
                    lblDiag.Font = new Font("Segoe UI", 10F);

                    _diagnosis.Location = new Point(150, 220);
                    _diagnosis.MaximumSize = new Size(900, 60);
                    _diagnosis.Size = new Size(370, 60);
                    _diagnosis.Font = new Font("Segoe UI", 10F);

                    // Medications
                    lblMeds.Location = new Point(20, 320);
                    lblMeds.Font = new Font("Segoe UI", 10F, FontStyle.Bold);

                    _medSearch.Location = new Point(260, 318);
                    _medSearch.Size = new Size(250, 28);
                    _medSearch.Font = new Font("Segoe UI", 10F);

                    _meds.Location = new Point(20, 350);
                    _meds.MaximumSize = new Size(900, 180);
                    _meds.Size = new Size(500, 180);
                    _meds.Font = new Font("Segoe UI", 10F);

                    // Vaccination checkbox
                    _markVaccinated.Location = new Point(20, 540);
                    _markVaccinated.Font = new Font("Segoe UI", 10F);
                    _markVaccinated.AutoSize = true;
                    _markVaccinated.BringToFront();

                    // Total
                    _totalLabel.Location = new Point(20, 585);
                    _totalLabel.Font = new Font("Segoe UI", 12F, FontStyle.Bold);

                    // Vaccination alert
                    _vaccinationAlert.Location = new Point(20, 615);
                    _vaccinationAlert.Font = new Font("Segoe UI", 10F, FontStyle.Bold);

                    // Buttons
                    save.Size = new Size(120, 45);
                    cancel.Size = new Size(120, 45);

                    save.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
                    cancel.Font = new Font("Segoe UI", 10F);

                    save.Location = new Point(
                        (ClientSize.Width / 2) - 140,
                        645
                    );

                    cancel.Location = new Point(
                        (ClientSize.Width / 2) + 20,
                        645
                    );
                }

                // מרכז המסך רק במצב גדול
                if (ClientSize.Width > 1200)
                {
                    contentPanel.Location = new Point(
                        (ClientSize.Width - contentPanel.Width) / 2,
                        175
                    );
                    bigHeader.BringToFront();
                }
                else
                {
                    // במסך קטן הכותרת הגדולה מוסתרת
                    bigHeader.Visible = false;

                    // מצב רגיל - מיקום רגיל
                    contentPanel.Location = new Point(0, 0);
                }
            };
            // מפעיל פעם אחת את קוד ה־Resize מיד בפתיחת החלון
            OnResize(EventArgs.Empty);
        }



        private PictureBox CreateDecorativeAnimalPicture(
                string fileName,
                Point location,
                Size size)
        {
            var picture = new PictureBox
            {
                Location = location,
                Size = size,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.Transparent
            };

            string imagePath = Path.Combine(
                Application.StartupPath,
                "Images",
                fileName);

            if (File.Exists(imagePath))
            {
                picture.Image = Image.FromFile(imagePath);
            }

            return picture;
        }



        // מיישר לפי התו שהמשתמש מקליד,
        // בלי להתערב ברווחים ובלי לשנות RightToLeft
        private void UpdateTextAlignmentByKey(
            RichTextBox box,
            char typedChar,
            ref bool currentlyHebrewAligned)
        {
            // אם זה placeholder אפור, לא עושים כלום
            if (box.ForeColor == Color.Gray)
                return;

            // אם זה רווח, Enter, Backspace וכו' — לא עושים כלום
            if (char.IsWhiteSpace(typedChar) || char.IsControl(typedChar))
                return;

            bool isHebrew =
                typedChar >= '\u0590' && typedChar <= '\u05FF';

            bool isEnglish =
                (typedChar >= 'A' && typedChar <= 'Z') ||
                (typedChar >= 'a' && typedChar <= 'z');

            // אם זה לא עברית ולא אנגלית, לא משנים יישור
            if (!isHebrew && !isEnglish)
                return;

            // אם הכיוון לא השתנה — לא נוגעים
            if (isHebrew == currentlyHebrewAligned)
                return;

            currentlyHebrewAligned = isHebrew;

            int selectionStart = box.SelectionStart;
            int selectionLength = box.SelectionLength;

            box.SelectionAlignment = isHebrew
                ? HorizontalAlignment.Right
                : HorizontalAlignment.Left;

            box.SelectionStart = Math.Min(selectionStart, box.TextLength);
            box.SelectionLength = selectionLength;
        }



        // מוסיף טקסט רמז ל־RichTextBox
        private void AddPlaceholder(RichTextBox box, string placeholderText)
        {
            box.Text = placeholderText;
            box.ForeColor = Color.Gray;

            box.GotFocus += (_, _) =>
            {
                if (box.ForeColor == Color.Gray && box.Text == placeholderText)
                {
                    box.Text = "";
                    box.ForeColor = Color.Black;
                }
            };

            box.LostFocus += (_, _) =>
            {
                if (string.IsNullOrWhiteSpace(box.Text))
                {
                    box.Text = placeholderText;
                    box.ForeColor = Color.Gray;
                    box.SelectionAlignment = HorizontalAlignment.Left;
                }
            };
        }



        /// <summary>
        /// בודקת האם החיה שנבחרה צריכה חיסון שנתי,
        /// ומציגה התראה מתאימה למסך.
        /// </summary>
        private void UpdateVaccinationAlert()
        {
            // אם אין חיה שנבחרה
            if (_animal.SelectedItem == null)
            {
                _vaccinationAlert.Text = "";
                return;
            }

            // קבלת הטקסט של החיה
            var animalText = _animal.SelectedItem.ToString() ?? "";

            // מציאת מספר השבב
            var chipStart = animalText.IndexOf("chip ", StringComparison.Ordinal) + 5;
            var chipEnd = animalText.IndexOf(')', chipStart);

            // חילוץ מספר השבב
            var chip = animalText.Substring(chipStart, chipEnd - chipStart);

            // חיפוש החיה במערכת
            var animal = _store.Animals
                .FirstOrDefault(a => a.ChipNumber == chip);

            // אם החיה לא נמצאה
            if (animal == null)
            {
                _vaccinationAlert.Text = "";
                return;
            }

            // בדיקה האם החיה צריכה חיסון
            bool needsVaccination =
                animal.LastVaccinationDate == null ||
                (DateTime.Now - animal.LastVaccinationDate.Value).TotalDays > 365;

            // הצגת ההתראה
            _vaccinationAlert.Text = needsVaccination
                ? "⚠ Last vaccination was over a year ago."
                : "";
        }


        // חיפוש חיה לפי מספר צ'יפ
        private void OnSearchChip(object? sender, EventArgs e)
        {
            // חלון חיפוש חיה
            var searchForm = new Form
            {
                Text = "Search Animal",
                Size = new Size(420, 220),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.White,
                Font = new Font("Segoe UI", 10F)
            };

            // כותרת
            var lbl = new Label
            {
                Text = "Enter animal chip number:",
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(70, 25)
            };

            // שדה צ'יפ
            var txtChip = new TextBox
            {
                Size = new Size(240, 30),
                Location = new Point(85, 70),
                Font = new Font("Segoe UI", 12F),
                TextAlign = HorizontalAlignment.Center
            };

            // כפתור אישור
            var okBtn = new Button
            {
                Text = "OK",
                Size = new Size(110, 40),
                Location = new Point(70, 130),
                BackColor = Color.FromArgb(40, 90, 150),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.OK
            };

            okBtn.FlatAppearance.BorderSize = 0;

            // כפתור ביטול
            var cancelBtn = new Button
            {
                Text = "Cancel",
                Size = new Size(110, 40),
                Location = new Point(220, 130),
                BackColor = Color.LightGray,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel
            };

            cancelBtn.FlatAppearance.BorderSize = 0;

            // הוספה לחלון
            searchForm.Controls.Add(lbl);
            searchForm.Controls.Add(txtChip);
            searchForm.Controls.Add(okBtn);
            searchForm.Controls.Add(cancelBtn);

            // הגדרת Enter / Escape
            searchForm.AcceptButton = okBtn;
            searchForm.CancelButton = cancelBtn;

            // פתיחת החלון
            if (searchForm.ShowDialog(this) != DialogResult.OK)
                return;

            // הצ'יפ שהוקלד
            string chip = txtChip.Text.Trim();

            // אם המשתמש ביטל
            if (string.IsNullOrWhiteSpace(chip))
                return;


            // חיפוש החיה ברשימה לפי צ'יפ מלא בלבד
            for (int i = 0; i < _animal.Items.Count; i++)
            {
                var itemText = _animal.Items[i]?.ToString() ?? "";

                // התאמה מלאה בלבד:
                // לדוגמה, חיפוש 567 לא יתפוס בטעות chip 5678
                if (itemText.Contains($"(chip {chip})", StringComparison.Ordinal))
                {
                    _animal.SelectedIndex = i;

                    MessageBox.Show(
                        "Animal found successfully.",
                        "Search",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    return;
                }
            }

            // אם לא נמצאה חיה
            MessageBox.Show(
                "No animal found with this chip number.",
                "Not found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }




        /// <summary>
        /// פותחת חלון לבחירת כמות לתרופה.
        /// מחזירה null אם המשתמש ביטל.
        /// </summary>
        private int? AskMedicationQuantity(string medicationName, int maxStock)
        {
            // חלון בחירת כמות
            using var quantityForm = new Form
            {
                Text = "Medication Quantity",
                Size = new Size(480, 300),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(245, 248, 252),
                Font = new Font("Segoe UI", 10F)
            };

            // כרטיס פנימי לבן
            var card = new Panel
            {
                BackColor = Color.White,
                Location = new Point(20, 20),
                Size = new Size(425, 210),
                BorderStyle = BorderStyle.None
            };

            // כותרת
            var title = new Label
            {
                Text = $"Choose quantity for {medicationName}",
                Font = new Font("Segoe UI", 13F, FontStyle.Bold),
                ForeColor = Color.FromArgb(40, 90, 150),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(20, 20),
                Size = new Size(385, 35)
            };

            // טקסט מלאי
            var stockLabel = new Label
            {
                Text = $"Available stock: {maxStock}",
                Font = new Font("Segoe UI", 10.5F),
                ForeColor = Color.FromArgb(60, 60, 60),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(20, 60),
                Size = new Size(385, 28)
            };

            // תווית כמות
            var quantityLabel = new Label
            {
                Text = "Quantity:",
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(120, 105)
            };

            // בחירת כמות
            var quantityBox = new NumericUpDown
            {
                Minimum = 1,
                Maximum = Math.Max(1, maxStock),
                Value = 1,
                Location = new Point(210, 101),
                Size = new Size(90, 30),
                Font = new Font("Segoe UI", 11F),
                TextAlign = HorizontalAlignment.Center
            };

            // כפתור אישור
            var okBtn = new Button
            {
                Text = "OK",
                Size = new Size(120, 42),
                Location = new Point(75, 155),
                BackColor = Color.FromArgb(40, 90, 150),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };

            okBtn.FlatAppearance.BorderSize = 0;

            // כפתור ביטול
            var cancelBtn = new Button
            {
                Text = "Cancel",
                Size = new Size(120, 42),
                Location = new Point(225, 155),
                BackColor = Color.FromArgb(220, 220, 220),
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F),
                DialogResult = DialogResult.Cancel
            };

            cancelBtn.FlatAppearance.BorderSize = 0;

            // הוספת הרכיבים לכרטיס
            card.Controls.Add(title);
            card.Controls.Add(stockLabel);
            card.Controls.Add(quantityLabel);
            card.Controls.Add(quantityBox);
            card.Controls.Add(okBtn);
            card.Controls.Add(cancelBtn);

            // הוספת הכרטיס לחלון
            quantityForm.Controls.Add(card);

            // הגדרת Enter / Escape
            quantityForm.AcceptButton = okBtn;
            quantityForm.CancelButton = cancelBtn;

            // פתיחת החלון
            if (quantityForm.ShowDialog(this) != DialogResult.OK)
                return null;

            return (int)quantityBox.Value;
        }



        /// <summary>
        /// טוענת את רשימת התרופות לתוך ה־CheckedListBox.
        /// אם המשתמש הקליד טקסט בחיפוש, מוצגות רק תרופות שמתאימות לחיפוש.
        /// </summary>
        private void LoadMedicationList(string filter = "")
        {
            // סימון שהרשימה נטענת עכשיו
            _loadingMedicationList = true;

            // ניקוי הרשימה הקיימת
            _meds.Items.Clear();

            // ניקוי טקסט החיפוש
            string searchText = filter.Trim().ToLower();

            // מעבר על כל התרופות במערכת
            foreach (var medication in _store.Medications)
            {
                // אם יש חיפוש והתרופה לא מתאימה, מדלגים עליה
                if (!string.IsNullOrWhiteSpace(searchText) &&
                    !medication.Name.ToLower().Contains(searchText))
                {
                    continue;
                }

                // בדיקה אם התרופה כבר נבחרה ומה הכמות שלה
                bool isSelected = _selectedMedicationQuantities.ContainsKey(medication.Name);
                int selectedQuantity = isSelected ? _selectedMedicationQuantities[medication.Name] : 0;

                // יצירת טקסט להצגה
                string displayText = medication.StockQuantity <= 0
                    ? $"{medication.Name} — OUT OF STOCK"
                    : isSelected
                        ? $"{medication.Name} — ₪{medication.Price:0.##} (Stock: {medication.StockQuantity}) [Qty: {selectedQuantity}]"
                        : $"{medication.Name} — ₪{medication.Price:0.##} (Stock: {medication.StockQuantity})";

                // הוספת התרופה לרשימה
                int index = _meds.Items.Add(displayText);

                // אם התרופה כבר נבחרה לפני החיפוש, מחזירים לה סימון
                if (isSelected)
                {
                    _meds.SetItemChecked(index, true);
                }
            }

            // סיום טעינה
            _loadingMedicationList = false;
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

            // מעבר על כל התרופות שסומנו ועל הכמות שלהן
            foreach (var medicationQuantity in _selectedMedicationQuantities)
            {
                var medicationName = medicationQuantity.Key;
                var quantity = medicationQuantity.Value;

                var medication = _store.Medications
                    .FirstOrDefault(m => m.Name == medicationName);

                if (medication != null)
                {
                    total += medication.Price * quantity;
                }
            }

            // עדכון תצוגת המחיר למשתמש
            _totalLabel.Text =
                $"Total: ₪{total:0.##} (base ₪{_basePrice:0.##} + meds ₪{total - _basePrice:0.##})";
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
            if (string.IsNullOrWhiteSpace(_reason.Text) ||
                _reason.ForeColor == Color.Gray)
            { Tell("Reason is required."); return; }

            // בדיקת חובה: אבחנה
            if (string.IsNullOrWhiteSpace(_diagnosis.Text) ||
                _diagnosis.ForeColor == Color.Gray)
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

            // מעבר על כל התרופות שסומנו ועל הכמות שלהן
            foreach (var medicationQuantity in _selectedMedicationQuantities)
            {
                // שם התרופה
                var medicationName = medicationQuantity.Key;

                // הכמות שניתנה
                var quantity = medicationQuantity.Value;

                // שמירת התרופה לרשימת התרופות של הביקור
                // אם הכמות היא 1, נשמור רק את השם.
                // אם הכמות גדולה מ־1, נשמור עם xQuantity.
                meds.Add(quantity == 1
                    ? medicationName
                    : $"{medicationName} x{quantity}");

                // חיפוש התרופה במערכת
                var medication = _store.Medications
                    .FirstOrDefault(m => m.Name == medicationName);

                // אם התרופה נמצאה, מוסיפים מחיר כפול כמות
                if (medication != null)
                {
                    medTotal += medication.Price * quantity;
                }
            }

            // שמירת כמויות התרופות כדי שהמסך הראשי יוכל להוריד מהמלאי
            MedicationQuantities = new Dictionary<string, int>(_selectedMedicationQuantities);

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