using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;
using ClinicVets.Models;

namespace ClinicVets.Forms;

public class LoginForm : Form
{
    private readonly ExcelDataStore _store;
    private readonly TextBox _usernameBox = new();
    private readonly TextBox _passwordBox = new();

    public LoginForm(ExcelDataStore store)
    {
        _store = store;
        BuildUi();
    }

    private void BuildUi()
    {
        Text = "ClinicVets — Login";
        Size = new Size(440, 360);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(245, 248, 252);
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var title = new Label
        {
            Text = "🐾 ClinicVets",
            Font = new Font("Segoe UI", 22F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            AutoSize = true,
            Location = new Point(120, 20),
        };

        var subtitle = new Label
        {
            Text = "Veterinary Clinic Management",
            Font = new Font("Segoe UI", 10F, FontStyle.Italic),
            ForeColor = Color.Gray,
            AutoSize = true,
            Location = new Point(120, 70),
        };

        var lblUser = new Label { Text = "Username:", Location = new Point(40, 120), AutoSize = true };
        _usernameBox.Location = new Point(150, 117);
        _usernameBox.Size = new Size(220, 25);

        var lblPass = new Label { Text = "Password:", Location = new Point(40, 160), AutoSize = true };
        _passwordBox.Location = new Point(150, 157);
        _passwordBox.Size = new Size(220, 25);
        _passwordBox.UseSystemPasswordChar = true;

        var loginBtn = new Button
        {
            Text = "Login",
            Location = new Point(150, 210),
            Size = new Size(100, 36),
            BackColor = Color.FromArgb(40, 120, 200),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
        };
        loginBtn.FlatAppearance.BorderSize = 0;
        loginBtn.Click += OnLoginClick;

        var registerBtn = new Button
        {
            Text = "Register",
            Location = new Point(270, 210),
            Size = new Size(100, 36),
            BackColor = Color.FromArgb(220, 220, 220),
            FlatStyle = FlatStyle.Flat,
        };
        registerBtn.FlatAppearance.BorderSize = 0;
        registerBtn.Click += OnRegisterClick;

        var hint = new Label
        {
            Text = "New employees can register a clinic account.",
            Font = new Font("Segoe UI", 8F, FontStyle.Italic),
            ForeColor = Color.Gray,
            AutoSize = true,
            Location = new Point(80, 270),
        };

        Controls.AddRange(new Control[] { title, subtitle, lblUser, _usernameBox, lblPass, _passwordBox, loginBtn, registerBtn, hint });
        AcceptButton = loginBtn;
    }

    public Employee? AuthenticatedEmployee { get; private set; }

    private void OnLoginClick(object? sender, EventArgs e)
    {
        var user = _usernameBox.Text.Trim();
        var pass = _passwordBox.Text;
        var match = _store.Employees.FirstOrDefault(emp => emp.Username == user && emp.Password == pass);
        if (match == null)
        {
            MessageBox.Show("Invalid username or password.", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            _passwordBox.Clear();
            _passwordBox.Focus();
            return;
        }
        AuthenticatedEmployee = match;
        DialogResult = DialogResult.OK;
        Close();
    }

    private void OnRegisterClick(object? sender, EventArgs e)
    {
        using var register = new RegisterForm(_store);
        register.ShowDialog(this);
    }
}
