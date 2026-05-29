using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;

namespace ClinicVets.Forms;

public class ManageSpeciesDialog : Form
{
    private readonly ExcelDataStore _store;
    private readonly ListBox _list = new();
    private readonly TextBox _newName = new();

    public ManageSpeciesDialog(ExcelDataStore store)
    {
        _store = store;
        BuildUi();
        RefreshList();
    }

    private void BuildUi()
    {
        Text = "Manage Species Catalog";
        Size = new Size(420, 420);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(245, 248, 252);
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimumSize = new Size(420, 420);

        var title = new Label
        {
            Text = "🐾 Species",
            Font = new Font("Segoe UI", 14F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            AutoSize = true,
            Location = new Point(20, 15),
        };
        Controls.Add(title);

        _list.Location = new Point(20, 55);
        _list.Size = new Size(360, 220);
        Controls.Add(_list);

        _newName.Location = new Point(20, 290);
        _newName.Size = new Size(220, 25);
        Controls.Add(_newName);

        var addBtn = UiHelpers.MakeButton("Add", new Point(250, 288), Color.FromArgb(40, 160, 90), size: new Size(60, 28));
        addBtn.Click += OnAdd;
        Controls.Add(addBtn);

        var removeBtn = UiHelpers.MakeButton("Remove", new Point(320, 288), Color.FromArgb(200, 80, 80), size: new Size(70, 28));
        removeBtn.Click += OnRemove;
        Controls.Add(removeBtn);

        var closeBtn = UiHelpers.MakeButton("Close", new Point(290, 335), Color.FromArgb(120, 130, 150), size: new Size(100, 30));
        closeBtn.Click += (_, _) => Close();
        Controls.Add(closeBtn);

        UiHelpers.EnableProportionalScaling(this);
    }

    private void RefreshList()
    {
        _list.Items.Clear();
        foreach (var s in _store.Species) _list.Items.Add(s);
    }

    private void OnAdd(object? sender, EventArgs e)
    {
        if (!_store.TryAddSpecies(_newName.Text, out var err))
        {
            MessageBox.Show(err, "Cannot add", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        _newName.Clear();
        RefreshList();
    }

    private void OnRemove(object? sender, EventArgs e)
    {
        if (_list.SelectedItem is not string name)
        {
            MessageBox.Show("Select a species to remove.", "No selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        if (MessageBox.Show($"Remove species '{name}'?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;
        _store.RemoveSpecies(name);
        RefreshList();
    }
}
