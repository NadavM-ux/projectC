using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Models;
using ClinicVets.Validation;

namespace ClinicVets.Forms;

public class AddAnimalDialog : Form
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

    public AddAnimalDialog(IEnumerable<Customer> customers, IEnumerable<string> speciesCatalog)
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

        _species.DropDownStyle = ComboBoxStyle.DropDownList;
        foreach (var s in speciesCatalog) _species.Items.Add(s);
        if (_species.Items.Count > 0) _species.SelectedIndex = 0;
        AddRow("Species:", _species, 100);

        _weight.DecimalPlaces = 1;
        _weight.Minimum = 0.1m;
        _weight.Maximum = 100m;
        _weight.Increment = 0.5m;
        _weight.Value = 5m;
        AddRow("Weight (0.1-100):", _weight, 140);

        var now = DateTime.Now;
        _dob.Format = DateTimePickerFormat.Short;
        _dob.MinDate = new DateTime(2000, 1, 1);
        _dob.MaxDate = now;
        _dob.Value = now;
        AddRow("Date of birth:", _dob, 180);

        _owner.DropDownStyle = ComboBoxStyle.DropDownList;
        _owner.FormattingEnabled = true;
        _owner.Format += (_, e) =>
        {
            if (e.ListItem is Customer c) e.Value = $"{c.FullName} — {c.NationalId}";
        };
        _owner.DataSource = customers.ToList();
        AddRow("Owner:", _owner, 220);

        _hasVaccine.Text = "Has prior vaccination on:";
        _hasVaccine.Location = new Point(20, 263);
        _hasVaccine.AutoSize = true;
        _hasVaccine.CheckedChanged += (_, _) => _vaccine.Enabled = _hasVaccine.Checked;
        _vaccine.Location = new Point(180, 260);
        _vaccine.Size = new Size(220, 25);
        _vaccine.Format = DateTimePickerFormat.Short;
        _vaccine.MaxDate = DateTime.Now;
        _vaccine.Enabled = false;
        Controls.Add(_hasVaccine);
        Controls.Add(_vaccine);

        var save = UiHelpers.MakeButton("Save", new Point(200, 360), Color.FromArgb(40, 160, 90), size: new Size(90, 32));
        save.Click += OnSave;

        var cancel = UiHelpers.MakeButton("Cancel", new Point(310, 360), Color.FromArgb(220, 220, 220), Color.Black, new Size(90, 32));
        cancel.Click += (_, _) => Close();

        Controls.Add(save);
        Controls.Add(cancel);
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
        if (string.IsNullOrWhiteSpace(_chip.Text)) { Tell("Chip number is required."); return; }
        if (!Validators.IsLettersOnly(_name.Text.Trim(), out var err, "Name")) { Tell(err); return; }
        if (!Validators.IsValidWeight((double)_weight.Value, out err)) { Tell(err); return; }
        if (!Validators.IsValidBirthDate(_dob.Value.Date, out err)) { Tell(err); return; }
        if (_species.SelectedItem == null) { Tell("Choose a species."); return; }
        if (_owner.SelectedItem is not Customer owner) { Tell("Choose an owner."); return; }

        Result = new Animal
        {
            ChipNumber = _chip.Text.Trim(),
            Name = _name.Text.Trim(),
            Species = _species.SelectedItem.ToString() ?? "",
            Weight = (double)_weight.Value,
            DateOfBirth = _dob.Value.Date,
            OwnerNationalId = owner.NationalId,
            LastVaccinationDate = _hasVaccine.Checked ? _vaccine.Value.Date : null,
        };
        DialogResult = DialogResult.OK;
        Close();
    }

    private static void Tell(string msg) =>
        MessageBox.Show(msg, "Validation error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
}
