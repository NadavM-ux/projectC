using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;
using ClinicVets.Models;

namespace ClinicVets.Forms;

public class AnimalsForm : AppForm
{
    private readonly ExcelDataStore _store;
    private readonly DataGridView _grid = new();
    private readonly TextBox _searchBox = new();
    private readonly ComboBox _searchBy = new();

    private readonly Label _title = new();
    private readonly Label _lblSearch = new();
    private Button _searchBtn = null!;
    private Button _clearBtn = null!;
    private Button _addBtn = null!;
    private Button _manageBtn = null!;
    private Button _backBtn = null!;

    public AnimalsForm(ExcelDataStore store)
    {
        _store = store;
        BuildUi();
        RefreshGrid(_store.Animals);
    }

    private void BuildUi()
    {
        Text = "Animal Patient Cards";
        Size = new Size(960, 600);
        MinimumSize = new Size(900, 560);
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(245, 248, 252);
        Font = new Font("Segoe UI", 10F);

        BuildHeader();
        BuildSearchBar();
        BuildGrid();
        BuildActionBar();

        // Responsive layout: positions are recomputed from the actual control
        // sizes on every resize, so growing the fonts at full screen never
        // makes the search controls overlap each other.
        Resize += (_, _) => LayoutControls();
        LayoutControls();
    }

    private void BuildHeader()
    {
        _title.Text = "🐶 Animals";
        _title.ForeColor = Color.FromArgb(40, 90, 150);
        _title.AutoSize = true;
        Controls.Add(_title);
    }

    private void BuildSearchBar()
    {
        _lblSearch.Text = "Search:";
        _lblSearch.AutoSize = true;
        Controls.Add(_lblSearch);

        _searchBox.BorderStyle = BorderStyle.FixedSingle;
        Controls.Add(_searchBox);

        _searchBy.DropDownStyle = ComboBoxStyle.DropDownList;
        _searchBy.Items.AddRange(new object[] { "Name", "Chip Number" });
        _searchBy.SelectedIndex = 0;
        Controls.Add(_searchBy);

        _searchBtn = UiHelpers.MakeButton("Search", Point.Empty, Color.FromArgb(40, 120, 200));
        _searchBtn.Click += OnSearch;
        Controls.Add(_searchBtn);

        _clearBtn = UiHelpers.MakeButton("Clear", Point.Empty, Color.FromArgb(220, 220, 220), Color.Black);
        _clearBtn.Click += (_, _) => { _searchBox.Clear(); RefreshGrid(_store.Animals); };
        Controls.Add(_clearBtn);
    }

    private void BuildGrid()
    {
        _grid.AutoGenerateColumns = false;
        _grid.AllowUserToAddRows = false;
        _grid.ReadOnly = true;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.MultiSelect = false;
        _grid.BackgroundColor = Color.White;
        _grid.RowHeadersVisible = false;
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        // Styled headers (also makes the header text scale with the chosen font).
        _grid.EnableHeadersVisualStyles = false;
        _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(40, 90, 150);
        _grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        _grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 248, 252);

        AddColumn("Chip", nameof(Animal.ChipNumber), 90);
        AddColumn("Name", nameof(Animal.Name), 130);
        AddColumn("Species", nameof(Animal.Species), 90);
        AddColumn("Weight (kg)", nameof(Animal.Weight), 90);
        AddColumn("DOB", nameof(Animal.DateOfBirth), 110, "dd-MM-yyyy");
        AddColumn("Owner ID", nameof(Animal.OwnerNationalId), 110);
        AddColumn("Last Vaccine", nameof(Animal.LastVaccinationDate), 110, "dd-MM-yyyy");

        UiHelpers.EnableDoubleBuffering(_grid);
        Controls.Add(_grid);
    }

    private void AddColumn(string header, string property, int width, string? format = null)
    {
        var col = new DataGridViewTextBoxColumn
        {
            HeaderText = header,
            DataPropertyName = property,
            FillWeight = width,
        };
        if (format != null) col.DefaultCellStyle.Format = format;
        _grid.Columns.Add(col);
    }

    private void BuildActionBar()
    {
        _addBtn = UiHelpers.MakeButton("➕ Add Animal", Point.Empty, Color.FromArgb(40, 160, 90));
        _addBtn.Click += OnAddAnimal;
        Controls.Add(_addBtn);

        _manageBtn = UiHelpers.MakeButton("🐾 Manage Species", Point.Empty, Color.FromArgb(80, 130, 180));
        _manageBtn.Click += OnManageSpecies;
        Controls.Add(_manageBtn);

        _backBtn = UiHelpers.MakeButton("← Main Menu", Point.Empty, Color.FromArgb(120, 130, 150));
        _backBtn.Click += (_, _) => Close();
        Controls.Add(_backBtn);
    }

    // Lays everything out for the current window size, scaling fonts when large
    // and spacing controls by their real widths so nothing overlaps.
    private void LayoutControls()
    {
        SuspendLayout();

        const int margin = 20;
        const int gap = 10;
        bool big = ClientSize.Width > 1100 || ClientSize.Height > 760;

        var ctrlFont = new Font("Segoe UI", big ? 14f : 10f);
        var btnFont = new Font("Segoe UI", big ? 12f : 9f, FontStyle.Bold);

        _title.Font = new Font("Segoe UI", big ? 24f : 16f, FontStyle.Bold);
        _title.Location = new Point(margin, margin);

        // ---- search row ----
        int rowH = big ? 42 : 30;
        int rowY = _title.Bottom + 15;

        _lblSearch.Font = ctrlFont;
        _searchBox.Font = ctrlFont;
        _searchBy.Font = ctrlFont;
        _searchBtn.Font = btnFont;
        _clearBtn.Font = btnFont;

        // Vertically centre a control inside the row (textbox/combo heights are
        // driven by their font, so we read their actual height).
        void PlaceCentered(Control c, int x, int? w = null)
        {
            if (w.HasValue) c.Width = w.Value;
            c.Location = new Point(x, rowY + (rowH - c.Height) / 2);
        }

        PlaceCentered(_lblSearch, margin);
        PlaceCentered(_searchBox, _lblSearch.Right + gap, big ? 340 : 200);
        PlaceCentered(_searchBy, _searchBox.Right + gap, big ? 200 : 140);

        int btnW = big ? 120 : 80;
        _searchBtn.Size = new Size(btnW, rowH);
        _searchBtn.Location = new Point(_searchBy.Right + gap, rowY);
        _clearBtn.Size = new Size(btnW, rowH);
        _clearBtn.Location = new Point(_searchBtn.Right + gap, rowY);

        // ---- action bar (bottom) ----
        int actH = big ? 46 : 36;
        int actY = ClientSize.Height - margin - actH;

        _addBtn.Font = btnFont;
        _addBtn.Size = new Size(big ? 210 : 160, actH);
        _addBtn.Location = new Point(margin, actY);

        _manageBtn.Font = btnFont;
        _manageBtn.Size = new Size(big ? 240 : 180, actH);
        _manageBtn.Location = new Point(_addBtn.Right + gap, actY);

        _backBtn.Font = btnFont;
        _backBtn.Size = new Size(big ? 190 : 140, actH);
        _backBtn.Location = new Point(_manageBtn.Right + gap, actY);

        // ---- grid fills the space between the search row and the action bar ----
        int gridTop = rowY + rowH + 15;
        _grid.Location = new Point(margin, gridTop);
        _grid.Size = new Size(
            ClientSize.Width - (margin * 2),
            System.Math.Max(120, actY - 15 - gridTop));

        _grid.DefaultCellStyle.Font = ctrlFont;
        _grid.ColumnHeadersDefaultCellStyle.Font =
            new Font("Segoe UI", big ? 14f : 10f, FontStyle.Bold);
        _grid.ColumnHeadersHeight = big ? 50 : 34;
        _grid.RowTemplate.Height = big ? 42 : 28;
        foreach (DataGridViewRow row in _grid.Rows) row.Height = big ? 42 : 28;

        ResumeLayout(true);
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
        if (_store.Species.Count == 0)
        {
            MessageBox.Show("The species catalog is empty. Add at least one species first.",
                "No species", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dlg = new AddAnimalDialog(_store.Customers, _store.Species);
        if (dlg.ShowDialog(this) != DialogResult.OK || dlg.Result == null) return;

        if (!_store.TryAddAnimal(dlg.Result, out var err))
        {
            MessageBox.Show(err, "Cannot add animal", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        RefreshGrid(_store.Animals);
    }

    private void OnManageSpecies(object? sender, EventArgs e)
    {
        using var dlg = new ManageSpeciesDialog(_store);
        dlg.ShowDialog(this);
    }
}
