using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;
using ClinicVets.Models;

namespace ClinicVets.Forms;

public class VisitsForm : Form
{
    private readonly ExcelDataStore _store;
    private readonly Employee _vet;
    private readonly DataGridView _grid = new();
    private readonly Label _alertLabel = new();

    private const decimal BaseVisitPrice = 150m;
    private const int VaccinationCadenceDays = 365;

    public VisitsForm(ExcelDataStore store, Employee vet)
    {
        _store = store;
        _vet = vet;
        BuildUi();
        RefreshGrid();
    }

    private void BuildUi()
    {
        Text = "Visits & Treatments";
        Size = new Size(960, 600);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(245, 248, 252);
        Font = new Font("Segoe UI", 10F);

        var title = new Label
        {
            Text = "🩺 Visits",
            Font = new Font("Segoe UI", 16F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            AutoSize = true,
            Location = new Point(20, 15),
        };
        Controls.Add(title);

        _alertLabel.Location = new Point(20, 55);
        _alertLabel.Size = new Size(900, 30);
        _alertLabel.ForeColor = Color.FromArgb(200, 70, 0);
        _alertLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
        Controls.Add(_alertLabel);
        ShowVaccinationAlerts();

        _grid.Location = new Point(20, 95);
        _grid.Size = new Size(900, 380);
        _grid.AutoGenerateColumns = false;
        _grid.AllowUserToAddRows = false;
        _grid.ReadOnly = true;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.BackgroundColor = Color.White;
        _grid.RowHeadersVisible = false;
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Date", DataPropertyName = nameof(Visit.VisitDateTime), Width = 130, DefaultCellStyle = { Format = "yyyy-MM-dd HH:mm" } });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Animal", DataPropertyName = nameof(Visit.AnimalChipNumber), Width = 100 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Vet", DataPropertyName = nameof(Visit.VetUsername), Width = 100 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Reason", DataPropertyName = nameof(Visit.Reason), Width = 180 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Diagnosis", DataPropertyName = nameof(Visit.Diagnosis), Width = 180 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Total (₪)", DataPropertyName = nameof(Visit.TotalPrice), Width = 100 });
        Controls.Add(_grid);

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
        addBtn.FlatAppearance.BorderSize = 0;
        addBtn.Click += OnNewVisit;
        Controls.Add(addBtn);

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
        backBtn.FlatAppearance.BorderSize = 0;
        backBtn.Click += (_, _) => Close();
        Controls.Add(backBtn);
    }

    private void RefreshGrid()
    {
        _grid.DataSource = null;
        _grid.DataSource = _store.Visits.OrderByDescending(v => v.VisitDateTime).ToList();
    }

    private void ShowVaccinationAlerts()
    {
        var due = _store.Animals
            .Where(a => a.LastVaccinationDate == null
                        || (DateTime.Now - a.LastVaccinationDate.Value).TotalDays > VaccinationCadenceDays)
            .ToList();
        _alertLabel.Text = due.Count == 0
            ? "✅ All animals are up to date on annual vaccinations."
            : $"⚠ {due.Count} animal(s) need an annual vaccination: " + string.Join(", ", due.Take(5).Select(a => a.Name));
    }

    private void OnNewVisit(object? sender, EventArgs e)
    {
        if (_store.Animals.Count == 0)
        {
            MessageBox.Show("There are no registered animals.", "No animals", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        using var dlg = new NewVisitDialog(_store, _vet, BaseVisitPrice);
        if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Result != null)
        {
            _store.Visits.Add(dlg.Result);
            if (dlg.UpdatedAnimal != null)
            {
                var idx = _store.Animals.FindIndex(a => a.ChipNumber == dlg.UpdatedAnimal.ChipNumber);
                if (idx >= 0) _store.Animals[idx] = dlg.UpdatedAnimal;
            }
            _store.Save();
            ShowVaccinationAlerts();
            RefreshGrid();
        }
    }

    private class NewVisitDialog : Form
    {
        private readonly ExcelDataStore _store;
        private readonly Employee _vet;
        private readonly decimal _basePrice;

        private readonly ComboBox _animal = new();
        private readonly DateTimePicker _date = new();
        private readonly TextBox _reason = new();
        private readonly TextBox _diagnosis = new();
        private readonly CheckedListBox _meds = new();
        private readonly Label _totalLabel = new();
        private readonly CheckBox _markVaccinated = new();

        public Visit? Result { get; private set; }
        public Animal? UpdatedAnimal { get; private set; }

        public NewVisitDialog(ExcelDataStore store, Employee vet, decimal basePrice)
        {
            _store = store;
            _vet = vet;
            _basePrice = basePrice;

            Text = "Open New Visit";
            Size = new Size(560, 640);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.FromArgb(245, 248, 252);
            Font = new Font("Segoe UI", 10F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;

            var lblAnimal = new Label { Text = "Animal:", Location = new Point(20, 23), AutoSize = true };
            _animal.Location = new Point(150, 20);
            _animal.Size = new Size(370, 25);
            _animal.DropDownStyle = ComboBoxStyle.DropDownList;
            foreach (var a in _store.Animals)
            {
                _animal.Items.Add($"{a.Name} (chip {a.ChipNumber}) — {a.Species}");
            }
            _animal.SelectedIndex = 0;
            Controls.Add(lblAnimal); Controls.Add(_animal);

            var lblDate = new Label { Text = "Date / time:", Location = new Point(20, 63), AutoSize = true };
            _date.Location = new Point(150, 60);
            _date.Size = new Size(370, 25);
            _date.Format = DateTimePickerFormat.Custom;
            _date.CustomFormat = "yyyy-MM-dd HH:mm";
            _date.Value = DateTime.Now;
            Controls.Add(lblDate); Controls.Add(_date);

            var lblReason = new Label { Text = "Reason for visit:", Location = new Point(20, 103), AutoSize = true };
            _reason.Location = new Point(150, 100);
            _reason.Size = new Size(370, 25);
            Controls.Add(lblReason); Controls.Add(_reason);

            var lblDiag = new Label { Text = "Diagnosis:", Location = new Point(20, 143), AutoSize = true };
            _diagnosis.Location = new Point(150, 140);
            _diagnosis.Size = new Size(370, 60);
            _diagnosis.Multiline = true;
            Controls.Add(lblDiag); Controls.Add(_diagnosis);

            var lblMeds = new Label { Text = "Medications administered:", Location = new Point(20, 213), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold) };
            _meds.Location = new Point(20, 240);
            _meds.Size = new Size(500, 180);
            _meds.CheckOnClick = true;
            foreach (var m in _store.Medications)
            {
                _meds.Items.Add($"{m.Name} — ₪{m.Price:0.##}");
            }
            _meds.ItemCheck += (_, _) => BeginInvoke(new Action(UpdateTotal));
            Controls.Add(lblMeds); Controls.Add(_meds);

            _markVaccinated.Text = "Mark this visit as today's vaccination";
            _markVaccinated.Location = new Point(20, 430);
            _markVaccinated.AutoSize = true;
            Controls.Add(_markVaccinated);

            _totalLabel.Text = $"Total: ₪{_basePrice:0.##} (base)";
            _totalLabel.Location = new Point(20, 470);
            _totalLabel.AutoSize = true;
            _totalLabel.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            _totalLabel.ForeColor = Color.FromArgb(40, 90, 150);
            Controls.Add(_totalLabel);

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
            save.FlatAppearance.BorderSize = 0;
            save.Click += OnSave;
            var cancel = new Button
            {
                Text = "Cancel",
                Location = new Point(400, 510),
                Size = new Size(110, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(220, 220, 220),
            };
            cancel.FlatAppearance.BorderSize = 0;
            cancel.Click += (_, _) => Close();
            Controls.Add(save); Controls.Add(cancel);
        }

        private void UpdateTotal()
        {
            decimal total = _basePrice;
            foreach (var item in _meds.CheckedItems)
            {
                var text = item?.ToString() ?? "";
                var parts = text.Split('₪');
                if (parts.Length == 2 && decimal.TryParse(parts[1], out var p))
                {
                    total += p;
                }
            }
            _totalLabel.Text = $"Total: ₪{total:0.##} (base ₪{_basePrice:0.##} + meds ₪{total - _basePrice:0.##})";
        }

        private void OnSave(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_reason.Text))
            { Tell("Reason is required."); return; }
            if (string.IsNullOrWhiteSpace(_diagnosis.Text))
            { Tell("Diagnosis is required."); return; }
            if (_animal.SelectedItem == null) { Tell("Choose an animal."); return; }

            var animalText = _animal.SelectedItem.ToString() ?? "";
            var chipStart = animalText.IndexOf("chip ", StringComparison.Ordinal) + 5;
            var chipEnd = animalText.IndexOf(')', chipStart);
            var chip = animalText.Substring(chipStart, chipEnd - chipStart);

            var meds = new List<string>();
            decimal medTotal = 0m;
            foreach (var item in _meds.CheckedItems)
            {
                var text = item?.ToString() ?? "";
                var dashIdx = text.IndexOf(" — ", StringComparison.Ordinal);
                var name = dashIdx > 0 ? text[..dashIdx] : text;
                meds.Add(name);
                var priceStr = text.Split('₪').LastOrDefault() ?? "0";
                if (decimal.TryParse(priceStr, out var p)) medTotal += p;
            }

            Result = new Visit
            {
                VisitId = Guid.NewGuid().ToString("N")[..8],
                AnimalChipNumber = chip,
                VetUsername = _vet.Username,
                VisitDateTime = _date.Value,
                Reason = _reason.Text.Trim(),
                Diagnosis = _diagnosis.Text.Trim(),
                MedicationsGiven = meds,
                BasePrice = _basePrice,
                TotalPrice = _basePrice + medTotal,
            };

            if (_markVaccinated.Checked)
            {
                var animal = _store.Animals.FirstOrDefault(a => a.ChipNumber == chip);
                if (animal != null)
                {
                    UpdatedAnimal = new Animal
                    {
                        ChipNumber = animal.ChipNumber,
                        Name = animal.Name,
                        Species = animal.Species,
                        Weight = animal.Weight,
                        DateOfBirth = animal.DateOfBirth,
                        OwnerNationalId = animal.OwnerNationalId,
                        LastVaccinationDate = _date.Value.Date,
                    };
                }
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private static void Tell(string msg) =>
            MessageBox.Show(msg, "Validation error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}
