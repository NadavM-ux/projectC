using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;
using ClinicVets.Models;

namespace ClinicVets.Forms;

public class MedicationsForm : Form
{
    private readonly ExcelDataStore _store;
    private readonly DataGridView _grid = new();

    public MedicationsForm(ExcelDataStore store)
    {
        _store = store;
        BuildUi();
        RefreshGrid();
    }

    private void BuildUi()
    {
        Text = "Medication Inventory";
        Size = new Size(700, 520);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(245, 248, 252);
        Font = new Font("Segoe UI", 10F);

        var title = new Label
        {
            Text = "💊 Medications",
            Font = new Font("Segoe UI", 16F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            AutoSize = true,
            Location = new Point(20, 15),
        };
        Controls.Add(title);

        _grid.Location = new Point(20, 60);
        _grid.Size = new Size(640, 340);
        _grid.AutoGenerateColumns = false;
        _grid.AllowUserToAddRows = false;
        _grid.ReadOnly = true;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.MultiSelect = false;
        _grid.BackgroundColor = Color.White;
        _grid.RowHeadersVisible = false;
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Name", DataPropertyName = nameof(Medication.Name), Width = 280 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Price (₪)", DataPropertyName = nameof(Medication.Price), Width = 130 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Stock", DataPropertyName = nameof(Medication.StockQuantity), Width = 130 });
        Controls.Add(_grid);

        var addBtn = MakeButton("➕ Add", new Point(20, 420), Color.FromArgb(40, 160, 90));
        addBtn.Click += OnAdd;
        var delBtn = MakeButton("🗑 Delete", new Point(130, 420), Color.FromArgb(200, 70, 70));
        delBtn.Click += OnDelete;
        var backBtn = MakeButton("← Main Menu", new Point(260, 420), Color.FromArgb(120, 130, 150));
        backBtn.Size = new Size(140, 36);
        backBtn.Click += (_, _) => Close();
        Controls.Add(addBtn); Controls.Add(delBtn); Controls.Add(backBtn);
    }

    private static Button MakeButton(string text, Point location, Color back)
    {
        var b = new Button
        {
            Text = text,
            Location = location,
            Size = new Size(100, 36),
            BackColor = back,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9F, FontStyle.Bold),
        };
        b.FlatAppearance.BorderSize = 0;
        return b;
    }

    private void RefreshGrid()
    {
        _grid.DataSource = null;
        _grid.DataSource = _store.Medications.ToList();
    }

    private void OnAdd(object? sender, EventArgs e)
    {
        using var dlg = new AddMedicationDialog();
        if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Result != null)
        {
            if (_store.Medications.Any(m => m.Name.Equals(dlg.Result.Name, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("A medication with this name already exists.", "Duplicate", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _store.Medications.Add(dlg.Result);
            _store.Save();
            RefreshGrid();
        }
    }

    private void OnDelete(object? sender, EventArgs e)
    {
        if (_grid.CurrentRow?.DataBoundItem is not Medication m) return;
        var confirm = MessageBox.Show($"Delete '{m.Name}'?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (confirm != DialogResult.Yes) return;
        _store.Medications.RemoveAll(med => med.Name == m.Name);
        _store.Save();
        RefreshGrid();
    }

    private class AddMedicationDialog : Form
    {
        private readonly TextBox _name = new();
        private readonly NumericUpDown _price = new();
        private readonly NumericUpDown _stock = new();

        public Medication? Result { get; private set; }

        public AddMedicationDialog()
        {
            Text = "Add Medication";
            Size = new Size(380, 260);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.FromArgb(245, 248, 252);
            Font = new Font("Segoe UI", 10F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;

            AddRow("Name:", _name, 20);

            var lblPrice = new Label { Text = "Price (₪):", Location = new Point(20, 63), AutoSize = true };
            _price.Location = new Point(150, 60);
            _price.Size = new Size(180, 25);
            _price.DecimalPlaces = 2;
            _price.Maximum = 100000;
            _price.Value = 50;
            Controls.Add(lblPrice); Controls.Add(_price);

            var lblStock = new Label { Text = "Stock quantity:", Location = new Point(20, 103), AutoSize = true };
            _stock.Location = new Point(150, 100);
            _stock.Size = new Size(180, 25);
            _stock.Maximum = 100000;
            _stock.Value = 10;
            Controls.Add(lblStock); Controls.Add(_stock);

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

            var cancel = new Button
            {
                Text = "Cancel",
                Location = new Point(240, 160),
                Size = new Size(90, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(220, 220, 220),
            };
            cancel.FlatAppearance.BorderSize = 0;
            cancel.Click += (_, _) => Close();

            Controls.Add(save); Controls.Add(cancel);
        }

        private void AddRow(string text, Control input, int y)
        {
            var lbl = new Label { Text = text, Location = new Point(20, y + 3), AutoSize = true };
            input.Location = new Point(150, y);
            input.Size = new Size(180, 25);
            Controls.Add(lbl);
            Controls.Add(input);
        }

        private void OnSave(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_name.Text))
            {
                MessageBox.Show("Name is required.", "Validation error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Result = new Medication
            {
                Name = _name.Text.Trim(),
                Price = _price.Value,
                StockQuantity = (int)_stock.Value,
            };
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
