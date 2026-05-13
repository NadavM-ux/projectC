using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;
using ClinicVets.Models;
using ClinicVets.Validation;

namespace ClinicVets.Forms;

public class RegisterForm : Form
{
    private readonly ExcelDataStore _store;
    private readonly TextBox _username = new();
    private readonly TextBox _password = new();
    private readonly TextBox _empNumber = new();
    private readonly TextBox _fullName = new();
    private readonly TextBox _email = new();
    private readonly TextBox _nationalId = new();
    private readonly ComboBox _role = new();

    public RegisterForm(ExcelDataStore store)
    {
        _store = store;
        BuildUi();
    }

    private void BuildUi()
    {
        Text = "Register New Employee";
        Size = new Size(500, 540);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(245, 248, 252);
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var title = new Label
        {
            Text = "Create Clinic Account",
            Font = new Font("Segoe UI", 16F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            AutoSize = true,
            Location = new Point(20, 15),
        };
        Controls.Add(title);

        AddField("Username (6-8, max 2 digits):", _username, 60);
        AddField("Password (8-10, letter+digit+special):", _password, 100);
        _password.UseSystemPasswordChar = true;
        AddField("Employee number (4 digits):", _empNumber, 140);
        AddField("Full name:", _fullName, 180);
        AddField("Email:", _email, 220);
        AddField("National ID (9 digits):", _nationalId, 260);

        var lblRole = new Label { Text = "Role:", Location = new Point(20, 303), AutoSize = true };
        _role.Location = new Point(230, 300);
        _role.Size = new Size(220, 25);
        _role.DropDownStyle = ComboBoxStyle.DropDownList;
        _role.Items.AddRange(new object[] { "Veterinarian", "Secretary" });
        _role.SelectedIndex = 0;
        Controls.Add(lblRole);
        Controls.Add(_role);

        var save = new Button
        {
            Text = "Register",
            Location = new Point(230, 360),
            Size = new Size(110, 36),
            BackColor = Color.FromArgb(40, 160, 90),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
        };
        save.FlatAppearance.BorderSize = 0;
        save.Click += OnSaveClick;

        var cancel = new Button
        {
            Text = "Cancel",
            Location = new Point(350, 360),
            Size = new Size(100, 36),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(220, 220, 220),
        };
        cancel.FlatAppearance.BorderSize = 0;
        cancel.Click += (_, _) => Close();

        Controls.Add(save);
        Controls.Add(cancel);

        var rules = new Label
        {
            Text = "Rules: username 6-8 chars (≤2 digits) • password 8-10 with letter, digit & one of !#$\n" +
                   "employee # exactly 4 digits • email must contain @ • national ID 9 digits + checksum",
            Font = new Font("Segoe UI", 8F, FontStyle.Italic),
            ForeColor = Color.Gray,
            AutoSize = true,
            Location = new Point(20, 420),
        };
        Controls.Add(rules);
    }

    private void AddField(string labelText, Control input, int y)
    {
        var lbl = new Label { Text = labelText, Location = new Point(20, y + 3), AutoSize = true };
        input.Location = new Point(230, y);
        input.Size = new Size(220, 25);
        Controls.Add(lbl);
        Controls.Add(input);
    }

    private void OnSaveClick(object? sender, EventArgs e)
    {
        if (!Validators.IsValidUsername(_username.Text.Trim(), out var err)) { Fail(err); return; }
        if (_store.Employees.Any(emp => emp.Username == _username.Text.Trim()))
        {
            Fail("Username already taken.");
            return;
        }
        if (!Validators.IsValidPassword(_password.Text, out err)) { Fail(err); return; }
        if (!Validators.IsValidEmployeeNumber(_empNumber.Text.Trim(), out err)) { Fail(err); return; }
        if (!Validators.IsLettersOnly(_fullName.Text.Trim(), out err, "Full name")) { Fail(err); return; }
        if (!Validators.IsValidEmail(_email.Text.Trim(), out err)) { Fail(err); return; }
        if (!Validators.IsValidNationalId(_nationalId.Text.Trim(), out err)) { Fail(err); return; }

        var role = _role.SelectedItem?.ToString() == "Veterinarian" ? Role.Veterinarian : Role.Secretary;

        _store.Employees.Add(new Employee
        {
            Username = _username.Text.Trim(),
            Password = _password.Text,
            EmployeeNumber = _empNumber.Text.Trim(),
            FullName = _fullName.Text.Trim(),
            Email = _email.Text.Trim(),
            NationalId = _nationalId.Text.Trim(),
            Role = role,
        });
        _store.Save();

        MessageBox.Show("Account created successfully. You can now log in.", "Registered", MessageBoxButtons.OK, MessageBoxIcon.Information);
        DialogResult = DialogResult.OK;
        Close();
    }

    private void Fail(string message)
    {
        MessageBox.Show(message, "Validation error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}
