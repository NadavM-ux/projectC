using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;

namespace ClinicVets.Forms;

public class ManageSpeciesDialog : AppForm
{
    private readonly ExcelDataStore _store;
    private readonly ListBox _list = new();
    private readonly TextBox _newName = new();

    private readonly Label _title = new();
    private Button _addBtn = null!;
    private Button _removeBtn = null!;
    private Button _closeBtn = null!;

    public ManageSpeciesDialog(ExcelDataStore store)
    {
        _store = store;
        BuildUi();
        RefreshList();
    }

    private void BuildUi()
    {
        Text = "Manage Species Catalog";
        Size = new Size(460, 460);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(245, 248, 252);
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimumSize = new Size(420, 420);

        _title.Text = "🐾 Species";
        _title.ForeColor = Color.FromArgb(40, 90, 150);
        _title.AutoSize = true;
        Controls.Add(_title);

        _list.BackColor = Color.White;
        _list.BorderStyle = BorderStyle.FixedSingle;
        _list.IntegralHeight = false; // allow the list to fill any height exactly
        Controls.Add(_list);

        _newName.BorderStyle = BorderStyle.FixedSingle;
        _newName.PlaceholderText = "New species name...";
        Controls.Add(_newName);

        _addBtn = UiHelpers.MakeButton("➕ Add", Point.Empty, Color.FromArgb(40, 160, 90));
        _addBtn.Click += OnAdd;
        Controls.Add(_addBtn);

        _removeBtn = UiHelpers.MakeButton("🗑 Remove", Point.Empty, Color.FromArgb(200, 80, 80));
        _removeBtn.Click += OnRemove;
        Controls.Add(_removeBtn);

        _closeBtn = UiHelpers.MakeButton("Close", Point.Empty, Color.FromArgb(120, 130, 150));
        _closeBtn.Click += (_, _) => Close();
        Controls.Add(_closeBtn);

        // Lay out responsively instead of uniformly scaling, so a maximized
        // window uses the whole area (no big empty margins) and looks clean.
        Resize += (_, _) => LayoutControls();
        LayoutControls();
    }

    // Positions everything based on the current client size. The list grows to
    // fill the window; the input box and buttons sit on a single bottom row.
    private void LayoutControls()
    {
        SuspendLayout();

        const int margin = 20;
        const int gap = 10;
        bool big = ClientSize.Width > 700 || ClientSize.Height > 700;

        _title.Font = new Font("Segoe UI", big ? 22F : 14F, FontStyle.Bold);
        _title.Location = new Point(margin, margin);

        int rowH = big ? 46 : 32;
        int rowY = ClientSize.Height - margin - rowH;

        var btnFont = new Font("Segoe UI", big ? 12F : 9F, FontStyle.Bold);

        _closeBtn.Font = btnFont;
        _closeBtn.Size = new Size(big ? 130 : 90, rowH);
        _closeBtn.Location = new Point(ClientSize.Width - margin - _closeBtn.Width, rowY);

        _removeBtn.Font = btnFont;
        _removeBtn.Size = new Size(big ? 130 : 95, rowH);
        _removeBtn.Location = new Point(_closeBtn.Left - gap - _removeBtn.Width, rowY);

        _addBtn.Font = btnFont;
        _addBtn.Size = new Size(big ? 110 : 75, rowH);
        _addBtn.Location = new Point(_removeBtn.Left - gap - _addBtn.Width, rowY);

        _newName.Font = new Font("Segoe UI", big ? 14F : 10F);
        int nameWidth = System.Math.Max(120, _addBtn.Left - gap - margin);
        _newName.Size = new Size(nameWidth, _newName.PreferredHeight);
        _newName.Location = new Point(margin, rowY + (rowH - _newName.Height) / 2);

        int listTop = _title.Bottom + 15;
        _list.Font = new Font("Segoe UI", big ? 14F : 10F);
        _list.Location = new Point(margin, listTop);
        _list.Size = new Size(
            ClientSize.Width - (margin * 2),
            System.Math.Max(60, rowY - gap - listTop));

        ResumeLayout(true);
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
