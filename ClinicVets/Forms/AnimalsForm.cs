using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;
using ClinicVets.Models;
using ClinicVets.Validation;

namespace ClinicVets.Forms;

public class AnimalsForm : Form
{
    private readonly ExcelDataStore _store;
    private readonly DataGridView _grid = new();
    private readonly TextBox _searchBox = new();
    private readonly ComboBox _searchBy = new();

    private static readonly string[] SpeciesCatalog = { "Dog", "Cat", "Reptile", "Bird" };

    public AnimalsForm(ExcelDataStore store)
    {
        _store = store;
        BuildUi();
        RefreshGrid(_store.Animals);
    }

    private void BuildUi()
    {
        Text = "Animal Patient Cards";
        Size = new Size(900, 560);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(245, 248, 252);
        Font = new Font("Segoe UI", 10F);

        var title = new Label
        {
            Text = "🐶 Animals",
            Font = new Font("Segoe UI", 16F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            AutoSize = true,
            Location = new Point(20, 15),
        };
        Controls.Add(title);

        var lblSearch = new Label { Text = "Search:", Location = new Point(20, 65), AutoSize = true };
        _searchBox.Location = new Point(85, 62);
        _searchBox.Size = new Size(200, 25);
        _searchBy.Location = new Point(295, 62);
        _searchBy.Size = new Size(140, 25);
        _searchBy.DropDownStyle = ComboBoxStyle.DropDownList;
        _searchBy.Items.AddRange(new object[] { "Name", "Chip Number" });
        _searchBy.SelectedIndex = 0;
        var searchBtn = MakeButton("Search", new Point(445, 60), Color.FromArgb(40, 120, 200));
        searchBtn.Click += OnSearch;
        var clearBtn = MakeButton("Clear", new Point(545, 60), Color.FromArgb(220, 220, 220), Color.Black);
        clearBtn.Click += (_, _) => { _searchBox.Clear(); RefreshGrid(_store.Animals); };
        Controls.AddRange(new Control[] { lblSearch, _searchBox, _searchBy, searchBtn, clearBtn });

        _grid.Location = new Point(20, 110);
        _grid.Size = new Size(840, 320);
        _grid.AutoGenerateColumns = false;
        _grid.AllowUserToAddRows = false;
        _grid.ReadOnly = true;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.BackgroundColor = Color.White;
        _grid.RowHeadersVisible = false;
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Chip", DataPropertyName = nameof(Animal.ChipNumber), Width = 90 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Name", DataPropertyName = nameof(Animal.Name), Width = 130 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Species", DataPropertyName = nameof(Animal.Species), Width = 90 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Weight (kg)", DataPropertyName = nameof(Animal.Weight), Width = 90 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "DOB", DataPropertyName = nameof(Animal.DateOfBirth), Width = 110, DefaultCellStyle = { Format = "yyyy-MM-dd" } });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Owner ID", DataPropertyName = nameof(Animal.OwnerNationalId), Width = 110 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Last Vaccine", DataPropertyName = nameof(Animal.LastVaccinationDate), Width = 110, DefaultCellStyle = { Format = "yyyy-MM-dd" } });
        Controls.Add(_grid);

        var addBtn = MakeButton("➕ Add Animal", new Point(20, 450), Color.FromArgb(40, 160, 90));
        addBtn.Size = new Size(160, 36);
        addBtn.Click += OnAddAnimal;
        Controls.Add(addBtn);

        var backBtn = MakeButton("← Main Menu", new Point(200, 450), Color.FromArgb(120, 130, 150));
        backBtn.Size = new Size(140, 36);
        backBtn.Click += (_, _) => Close();
        Controls.Add(backBtn);
    }

    private static Button MakeButton(string text, Point location, Color back, Color? fore = null)
    {
        var b = new Button
        {
            Text = text,
            Location = location,
            Size = new Size(90, 30),
            BackColor = back,
            ForeColor = fore ?? Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9F, FontStyle.Bold),
        };
        b.FlatAppearance.BorderSize = 0;
        return b;
    }

    private void OnSearch(object? sender, EventArgs e)
    {
        var query = _searchBox.Text.Trim();
        if (string.IsNullOrEmpty(query))
        {
            RefreshGrid(_store.Animals);
            return;
        }
        var by = _searchBy.SelectedItem?.ToString();
        IEnumerable<Animal> result = by == "Name"
            ? _store.Animals.Where(a => a.Name.Contains(query, StringComparison.OrdinalIgnoreCase))
            : _store.Animals.Where(a => a.ChipNumber.Contains(query));
        RefreshGrid(result);
    }

    private void RefreshGrid(IEnumerable<Animal> data)
    {
        _grid.DataSource = null;
        _grid.DataSource = data.ToList();
    }

    private void OnAddAnimal(object? sender, EventArgs e)
    {
        if (_store.Customers.Count == 0)
        {
            MessageBox.Show("There are no customers yet. A secretary must register a customer before adding animals.",
                "No customers", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        using var dlg = new AddAnimalDialog(_store.Customers);
        if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Result != null)
        {
            if (_store.Animals.Any(a => a.ChipNumber == dlg.Result.ChipNumber))
            {
                MessageBox.Show("Chip number already exists.", "Duplicate", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _store.Animals.Add(dlg.Result);
            _store.Save();
            RefreshGrid(_store.Animals);
        }
    }

    private class AddAnimalDialog : Form
    {
        private readonly TextBox _chip = new();
        private readonly TextBox _name = new();
        private readonly ComboBox _species = new();
        private readonly NumericUpDown _weight = new();
        private readonly DateTimePicker _dob = new();
        private readonly ComboBox _owner = new();
        private readonly DateTimePicker _vaccine = new();
        private readonly CheckBox _hasVaccine = new();

        public Animal? Result { get; private set; }

        public AddAnimalDialog(IEnumerable<Customer> customers)
        {
            Text = "Add Animal";
            Size = new Size(440, 480);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.FromArgb(245, 248, 252);
            Font = new Font("Segoe UI", 10F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;

            AddRow("Chip / serial #:", _chip, 20);
            AddRow("Name (letters):", _name, 60);

            var lblSpecies = new Label { Text = "Species:", Location = new Point(20, 103), AutoSize = true };
            _species.Location = new Point(180, 100);
            _species.Size = new Size(220, 25);
            _species.DropDownStyle = ComboBoxStyle.DropDownList;
            _species.Items.AddRange(SpeciesCatalog);
            _species.SelectedIndex = 0;
            Controls.Add(lblSpecies); Controls.Add(_species);

            var lblWeight = new Label { Text = "Weight (0.1-100):", Location = new Point(20, 143), AutoSize = true };
            _weight.Location = new Point(180, 140);
            _weight.Size = new Size(220, 25);
            _weight.DecimalPlaces = 1;
            _weight.Minimum = 0.1m;
            _weight.Maximum = 100m;
            _weight.Increment = 0.5m;
            _weight.Value = 5m;
            Controls.Add(lblWeight); Controls.Add(_weight);

            var lblDob = new Label { Text = "Date of birth:", Location = new Point(20, 183), AutoSize = true };
            var now = DateTime.Now;
            _dob.Location = new Point(180, 180);
            _dob.Size = new Size(220, 25);
            _dob.Format = DateTimePickerFormat.Short;
            _dob.MinDate = new DateTime(2000, 1, 1);
            _dob.MaxDate = now;
            _dob.Value = now;
            Controls.Add(lblDob); Controls.Add(_dob);

            var lblOwner = new Label { Text = "Owner:", Location = new Point(20, 223), AutoSize = true };
            _owner.Location = new Point(180, 220);
            _owner.Size = new Size(220, 25);
            _owner.DropDownStyle = ComboBoxStyle.DropDownList;
            foreach (var c in customers)
            {
                _owner.Items.Add($"{c.FullName} — {c.NationalId}");
            }
            if (_owner.Items.Count > 0) _owner.SelectedIndex = 0;
            Controls.Add(lblOwner); Controls.Add(_owner);

            _hasVaccine.Text = "Has prior vaccination on:";
            _hasVaccine.Location = new Point(20, 263);
            _hasVaccine.AutoSize = true;
            _hasVaccine.CheckedChanged += (_, _) => _vaccine.Enabled = _hasVaccine.Checked;
            _vaccine.Location = new Point(180, 260);
            _vaccine.Size = new Size(220, 25);
            _vaccine.Format = DateTimePickerFormat.Short;
            _vaccine.MaxDate = DateTime.Now;
            _vaccine.Enabled = false;
            Controls.Add(_hasVaccine); Controls.Add(_vaccine);

            var save = new Button
            {
                Text = "Save",
                Location = new Point(200, 360),
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
                Location = new Point(310, 360),
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
            input.Location = new Point(180, y);
            input.Size = new Size(220, 25);
            Controls.Add(lbl);
            Controls.Add(input);
        }

        private void OnSave(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_chip.Text))
            { Tell("Chip number is required."); return; }
            if (!Validators.IsLettersOnly(_name.Text.Trim(), out var err, "Name")) { Tell(err); return; }
            if (!Validators.IsValidWeight((double)_weight.Value, out err)) { Tell(err); return; }
            if (!Validators.IsValidBirthDate(_dob.Value.Date, out err)) { Tell(err); return; }
            if (_owner.SelectedItem == null) { Tell("Choose an owner."); return; }

            var ownerText = _owner.SelectedItem.ToString() ?? "";
            var ownerId = ownerText.Split('—').Last().Trim();

            Result = new Animal
            {
                ChipNumber = _chip.Text.Trim(),
                Name = _name.Text.Trim(),
                Species = _species.SelectedItem?.ToString() ?? "Dog",
                Weight = (double)_weight.Value,
                DateOfBirth = _dob.Value.Date,
                OwnerNationalId = ownerId,
                LastVaccinationDate = _hasVaccine.Checked ? _vaccine.Value.Date : null,
            };
            DialogResult = DialogResult.OK;
            Close();
        }

        private static void Tell(string msg) =>
            MessageBox.Show(msg, "Validation error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}
