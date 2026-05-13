using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;
using ClinicVets.Models;
using ClinicVets.Validation;

namespace ClinicVets.Forms;

public class CustomersForm : Form
{
    private readonly ExcelDataStore _store;
    private readonly DataGridView _grid = new();
    private readonly TextBox _searchBox = new();
    private readonly ComboBox _searchBy = new();
    private readonly ListBox _animalsList = new();
    private readonly Label _selectedCustomerLabel = new();

    public CustomersForm(ExcelDataStore store)
    {
        _store = store;
        BuildUi();
        RefreshGrid(_store.Customers);
    }

    private void BuildUi()
    {
        Text = "Customer Management";
        Size = new Size(900, 600);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(245, 248, 252);
        Font = new Font("Segoe UI", 10F);

        var title = new Label
        {
            Text = "👥 Customers",
            Font = new Font("Segoe UI", 16F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            AutoSize = true,
            Location = new Point(20, 15),
        };
        Controls.Add(title);

        var searchPanel = new Panel { Location = new Point(20, 55), Size = new Size(600, 40) };
        var lblSearch = new Label { Text = "Search:", Location = new Point(0, 10), AutoSize = true };
        _searchBox.Location = new Point(70, 7);
        _searchBox.Size = new Size(200, 25);
        _searchBy.Location = new Point(280, 7);
        _searchBy.Size = new Size(120, 25);
        _searchBy.DropDownStyle = ComboBoxStyle.DropDownList;
        _searchBy.Items.AddRange(new object[] { "National ID", "Phone" });
        _searchBy.SelectedIndex = 0;
        var searchBtn = MakeButton("Search", new Point(410, 5), Color.FromArgb(40, 120, 200));
        searchBtn.Click += OnSearch;
        var clearBtn = MakeButton("Clear", new Point(510, 5), Color.FromArgb(220, 220, 220), Color.Black);
        clearBtn.Click += (_, _) => { _searchBox.Clear(); RefreshGrid(_store.Customers); };
        searchPanel.Controls.AddRange(new Control[] { lblSearch, _searchBox, _searchBy, searchBtn, clearBtn });
        Controls.Add(searchPanel);

        _grid.Location = new Point(20, 110);
        _grid.Size = new Size(600, 250);
        _grid.AutoGenerateColumns = false;
        _grid.AllowUserToAddRows = false;
        _grid.ReadOnly = true;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.MultiSelect = false;
        _grid.BackgroundColor = Color.White;
        _grid.RowHeadersVisible = false;
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Full Name", DataPropertyName = nameof(Customer.FullName), Width = 180 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "National ID", DataPropertyName = nameof(Customer.NationalId), Width = 110 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Phone", DataPropertyName = nameof(Customer.Phone), Width = 120 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Email", DataPropertyName = nameof(Customer.Email), Width = 180 });
        _grid.SelectionChanged += OnSelectionChanged;
        Controls.Add(_grid);

        var addBtn = MakeButton("➕ Add Customer", new Point(20, 380), Color.FromArgb(40, 160, 90));
        addBtn.Size = new Size(160, 36);
        addBtn.Click += OnAddCustomer;
        Controls.Add(addBtn);

        var backBtn = MakeButton("← Main Menu", new Point(200, 380), Color.FromArgb(120, 130, 150));
        backBtn.Size = new Size(140, 36);
        backBtn.Click += (_, _) => Close();
        Controls.Add(backBtn);

        _selectedCustomerLabel.Text = "Animals owned by selected customer:";
        _selectedCustomerLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
        _selectedCustomerLabel.AutoSize = true;
        _selectedCustomerLabel.Location = new Point(640, 110);
        Controls.Add(_selectedCustomerLabel);

        _animalsList.Location = new Point(640, 140);
        _animalsList.Size = new Size(220, 380);
        _animalsList.BackColor = Color.White;
        Controls.Add(_animalsList);
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
            RefreshGrid(_store.Customers);
            return;
        }
        var by = _searchBy.SelectedItem?.ToString();
        IEnumerable<Customer> result = by == "National ID"
            ? _store.Customers.Where(c => c.NationalId.Contains(query))
            : _store.Customers.Where(c => c.Phone.Contains(query));
        RefreshGrid(result);
    }

    private void RefreshGrid(IEnumerable<Customer> data)
    {
        _grid.DataSource = null;
        _grid.DataSource = data.ToList();
    }

    private void OnSelectionChanged(object? sender, EventArgs e)
    {
        _animalsList.Items.Clear();
        if (_grid.CurrentRow?.DataBoundItem is not Customer customer) return;
        _selectedCustomerLabel.Text = $"Animals of {customer.FullName}:";
        var animals = _store.Animals.Where(a => a.OwnerNationalId == customer.NationalId);
        foreach (var a in animals)
        {
            _animalsList.Items.Add($"{a.Name} ({a.Species}) — chip {a.ChipNumber}");
        }
        if (_animalsList.Items.Count == 0)
        {
            _animalsList.Items.Add("(no animals on file)");
        }
    }

    private void OnAddCustomer(object? sender, EventArgs e)
    {
        using var dlg = new AddCustomerDialog();
        if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Result != null)
        {
            if (_store.Customers.Any(c => c.NationalId == dlg.Result.NationalId))
            {
                MessageBox.Show("A customer with this national ID already exists.", "Duplicate", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _store.Customers.Add(dlg.Result);
            _store.Save();
            RefreshGrid(_store.Customers);
        }
    }

    private class AddCustomerDialog : Form
    {
        private readonly TextBox _name = new();
        private readonly TextBox _id = new();
        private readonly TextBox _phone = new();
        private readonly TextBox _email = new();

        public Customer? Result { get; private set; }

        public AddCustomerDialog()
        {
            Text = "Add Customer";
            Size = new Size(420, 320);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.FromArgb(245, 248, 252);
            Font = new Font("Segoe UI", 10F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;

            AddRow("Full name (letters):", _name, 20);
            AddRow("National ID (9 digits):", _id, 60);
            AddRow("Phone:", _phone, 100);
            AddRow("Email:", _email, 140);

            var save = new Button
            {
                Text = "Save",
                Location = new Point(190, 200),
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
                Location = new Point(290, 200),
                Size = new Size(90, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(220, 220, 220),
            };
            cancel.FlatAppearance.BorderSize = 0;
            cancel.Click += (_, _) => Close();

            Controls.Add(save);
            Controls.Add(cancel);
        }

        private void AddRow(string text, Control input, int y)
        {
            var lbl = new Label { Text = text, Location = new Point(20, y + 3), AutoSize = true };
            input.Location = new Point(180, y);
            input.Size = new Size(200, 25);
            Controls.Add(lbl);
            Controls.Add(input);
        }

        private void OnSave(object? sender, EventArgs e)
        {
            if (!Validators.IsLettersOnly(_name.Text.Trim(), out var err, "Full name")) { Show(err); return; }
            if (!Validators.IsValidNationalId(_id.Text.Trim(), out err)) { Show(err); return; }
            if (!Validators.IsValidPhone(_phone.Text.Trim(), out err)) { Show(err); return; }
            if (!Validators.IsValidEmail(_email.Text.Trim(), out err)) { Show(err); return; }

            Result = new Customer
            {
                FullName = _name.Text.Trim(),
                NationalId = _id.Text.Trim(),
                Phone = _phone.Text.Trim(),
                Email = _email.Text.Trim(),
            };
            DialogResult = DialogResult.OK;
            Close();
        }

        private static void Show(string msg) =>
            MessageBox.Show(msg, "Validation error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}
