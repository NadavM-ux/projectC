using System;
using System.Drawing;
using System.IO;
using System.Linq;
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
        
        // Form size significantly increased to 800x600 to fully display the background artwork
        Size = new Size(800, 600); 
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        // --- Form Background Setup ---
        try
        {
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "login_bg.png");
            
            if (File.Exists(imagePath))
            {
                BackgroundImage = Image.FromFile(imagePath);
                BackgroundImageLayout = ImageLayout.Stretch;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error loading background image: " + ex.Message);
            BackColor = Color.FromArgb(245, 248, 252);
        }

        // --- Semi-transparent Card Panel ---
        // Resized slightly to 420x350 for better spacing inside a massive 800x600 viewport
        var cardPanel = new Panel
        {
            Size = new Size(420, 350),
            BackColor = Color.FromArgb(235, 255, 255, 255), 
        };
        
        // Automatically centers itself mathematically within the newly expanded client bounds
        cardPanel.Location = new Point(
            (this.ClientSize.Width - cardPanel.Width) / 2, 
            (this.ClientSize.Height - cardPanel.Height) / 2
        );

        // --- Controls Setup (Locations adjusted for the new 420x350 panel canvas) ---
        var title = new Label
        {
            Text = "🐾 ClinicVets",
            Font = new Font("Segoe UI", 24F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            BackColor = Color.Transparent, 
            AutoSize = true,
            Location = new Point(110, 25),
        };

        var subtitle = new Label
        {
            Text = "Veterinary Clinic Management",
            Font = new Font("Segoe UI", 10F, FontStyle.Italic),
            ForeColor = Color.DimGray,
            BackColor = Color.Transparent,
            AutoSize = true,
            Location = new Point(110, 75),
        };

        var lblUser = new Label 
        { 
            Text = "Username:", 
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            Location = new Point(40, 135), 
            AutoSize = true,
            BackColor = Color.Transparent
        };
        _usernameBox.Location = new Point(140, 132);
        _usernameBox.Size = new Size(240, 25);

        var lblPass = new Label 
        { 
            Text = "Password:", 
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            Location = new Point(40, 185), 
            AutoSize = true,
            BackColor = Color.Transparent
        };
        _passwordBox.Location = new Point(140, 182);
        _passwordBox.Size = new Size(240, 25);
        _passwordBox.UseSystemPasswordChar = true;

        var loginBtn = new Button
        {
            Text = "Login",
            Location = new Point(140, 235),
            Size = new Size(115, 36),
            BackColor = Color.FromArgb(40, 120, 200),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        loginBtn.FlatAppearance.BorderSize = 0;
        loginBtn.Click += OnLoginClick;

        var registerBtn = new Button
        {
            Text = "Register",
            Location = new Point(265, 235),
            Size = new Size(115, 36),
            BackColor = Color.FromArgb(220, 220, 220),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        registerBtn.FlatAppearance.BorderSize = 0;
        registerBtn.Click += OnRegisterClick;

        var hint = new Label
        {
            Text = "New employees can register a clinic account.",
            Font = new Font("Segoe UI", 8.5F, FontStyle.Italic),
            ForeColor = Color.DimGray,
            BackColor = Color.Transparent,
            AutoSize = true,
            Location = new Point(80, 295),
        };

        // Populate card sub-hierarchy
        cardPanel.Controls.AddRange(new Control[] { title, subtitle, lblUser, _usernameBox, lblPass, _passwordBox, loginBtn, registerBtn, hint });
        
        // Mount the master layout panel to the root form architecture
        Controls.Add(cardPanel);
        
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