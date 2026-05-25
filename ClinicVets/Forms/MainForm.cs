using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;
using ClinicVets.Models;

namespace ClinicVets.Forms;

public class MainForm : Form
{
    private readonly ExcelDataStore _store;
    private readonly Employee _user;

    public MainForm(ExcelDataStore store, Employee user)
    {
        _store = store;
        _user = user;
        BuildUi();
    }

    private void BuildUi()
    {
        Text = $"ClinicVets — {_user.FullName} ({_user.Role})";
        Size = new Size(720, 480);

        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        FormBorderStyle = FormBorderStyle.Sizable;
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(245, 248, 252);
        Font = new Font("Segoe UI", 10F);

        var header = new Panel
        {
            Dock = DockStyle.Top,
            Height = 80,
            BackColor = Color.FromArgb(40, 90, 150),
        };
        var hello = new Label
        {
            Text = $"🐾 Welcome, {_user.FullName}",
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 18F, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(20, 15),
        };
        var roleLabel = new Label
        {
            Text = _user.Role == Role.Veterinarian ? "Role: Veterinarian" : "Role: Secretary",
            ForeColor = Color.LightSkyBlue,
            Font = new Font("Segoe UI", 10F, FontStyle.Italic),
            AutoSize = true,
            Location = new Point(22, 50),
        };
        header.Controls.Add(hello);
        header.Controls.Add(roleLabel);
        Controls.Add(header);

        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 2,
            Padding = new Padding(40, 30, 40, 30),
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

        var customersBtn = MakeTile("👥 Customers", "Manage clients", _user.Role == Role.Secretary, OnCustomers);
        var animalsBtn = MakeTile("🐶 Animals", "Patient cards", true, OnAnimals);
        var visitsBtn = MakeTile("🩺 Visits", "Open & treat", _user.Role == Role.Veterinarian, OnVisits);
        var medsBtn = MakeTile("💊 Medications", "Inventory", _user.Role == Role.Veterinarian, OnMedications);

        grid.Controls.Add(customersBtn, 0, 0);
        grid.Controls.Add(animalsBtn, 1, 0);
        grid.Controls.Add(visitsBtn, 0, 1);
        grid.Controls.Add(medsBtn, 1, 1);

        Controls.Add(grid);

        var logout = new Button
        {
            Text = "Logout",
            Dock = DockStyle.Bottom,
            Height = 36,
            BackColor = Color.FromArgb(220, 220, 220),
            FlatStyle = FlatStyle.Flat,
        };
        logout.FlatAppearance.BorderSize = 0;
        logout.Click += (_, _) => { Application.Restart(); Environment.Exit(0); };
        Controls.Add(logout);
    }

    private Button MakeTile(string title, string subtitle, bool enabled, EventHandler onClick)
    {
        var btn = new Button
        {
            Text = $"{title}\n{subtitle}",
            Dock = DockStyle.Fill,
            FlatStyle = FlatStyle.Flat,
            BackColor = enabled ? Color.White : Color.FromArgb(235, 235, 235),
            ForeColor = enabled ? Color.FromArgb(40, 90, 150) : Color.Gray,
            Font = new Font("Segoe UI", 14F, FontStyle.Bold),
            Margin = new Padding(10),
            Enabled = enabled,
        };
        btn.FlatAppearance.BorderColor = Color.FromArgb(200, 210, 230);
        btn.FlatAppearance.BorderSize = 1;
        btn.Click += onClick;
        return btn;
    }

    private void OnCustomers(object? sender, EventArgs e)
    {
        using var f = new CustomersForm(_store);
        f.ShowDialog(this);
    }

    private void OnAnimals(object? sender, EventArgs e)
    {
        using var f = new AnimalsForm(_store);
        f.ShowDialog(this);
    }

    private void OnVisits(object? sender, EventArgs e)
    {
        using var f = new VisitsForm(_store, _user);
        f.ShowDialog(this);
    }

    private void OnMedications(object? sender, EventArgs e)
    {
        using var f = new MedicationsForm(_store);
        f.ShowDialog(this);
    }
}
