using System.Drawing;
using System.Windows.Forms;
using ClinicVets.Data;
using ClinicVets.Models;

namespace ClinicVets.Forms;

public class AnimalsForm : Form
{
    private readonly ExcelDataStore _store;
    private readonly DataGridView _grid = new();
    private readonly TextBox _searchBox = new();
    private readonly ComboBox _searchBy = new();

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
        MinimumSize = new Size(700, 460);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(245, 248, 252);
        Font = new Font("Segoe UI", 10F);

        BuildHeader();
        BuildSearchBar();
        BuildGrid();
        BuildActionBar();
    }

    private void BuildHeader()
    {
        var title = new Label
        {
            Text = "🐶 Animals",
            Font = new Font("Segoe UI", 16F, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 90, 150),
            AutoSize = true,
            Location = new Point(20, 15),
        };
        Controls.Add(title);
    }

    private void BuildSearchBar()
    {
        var lblSearch = new Label { Text = "Search:", Location = new Point(20, 65), AutoSize = true };
        _searchBox.Location = new Point(85, 62);
        _searchBox.Size = new Size(200, 25);
        _searchBy.Location = new Point(295, 62);
        _searchBy.Size = new Size(140, 25);
        _searchBy.DropDownStyle = ComboBoxStyle.DropDownList;
        _searchBy.Items.AddRange(new object[] { "Name", "Chip Number" });
        _searchBy.SelectedIndex = 0;

        var searchBtn = UiHelpers.MakeButton("Search", new Point(445, 60), Color.FromArgb(40, 120, 200));
        searchBtn.Click += OnSearch;

        var clearBtn = UiHelpers.MakeButton("Clear", new Point(545, 60), Color.FromArgb(220, 220, 220), Color.Black);
        clearBtn.Click += (_, _) => { _searchBox.Clear(); RefreshGrid(_store.Animals); };

        Controls.AddRange(new Control[] { lblSearch, _searchBox, _searchBy, searchBtn, clearBtn });
    }

    private void BuildGrid()
    {
        _grid.Location = new Point(20, 110);
        _grid.Size = new Size(840, 320);
        _grid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        _grid.AutoGenerateColumns = false;
        _grid.AllowUserToAddRows = false;
        _grid.ReadOnly = true;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.BackgroundColor = Color.White;
        _grid.RowHeadersVisible = false;
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        AddColumn("Chip", nameof(Animal.ChipNumber), 90);
        AddColumn("Name", nameof(Animal.Name), 130);
        AddColumn("Species", nameof(Animal.Species), 90);
        AddColumn("Weight (kg)", nameof(Animal.Weight), 90);
        AddColumn("DOB", nameof(Animal.DateOfBirth), 110, "yyyy-MM-dd");
        AddColumn("Owner ID", nameof(Animal.OwnerNationalId), 110);
        AddColumn("Last Vaccine", nameof(Animal.LastVaccinationDate), 110, "yyyy-MM-dd");

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
        var addBtn = UiHelpers.MakeButton("➕ Add Animal", new Point(20, 450), Color.FromArgb(40, 160, 90), size: new Size(160, 36));
        addBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        addBtn.Click += OnAddAnimal;
        Controls.Add(addBtn);

        var manageBtn = UiHelpers.MakeButton("🐾 Manage Species", new Point(200, 450), Color.FromArgb(80, 130, 180), size: new Size(180, 36));
        manageBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        manageBtn.Click += OnManageSpecies;
        Controls.Add(manageBtn);

        var backBtn = UiHelpers.MakeButton("← Main Menu", new Point(400, 450), Color.FromArgb(120, 130, 150), size: new Size(140, 36));
        backBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        backBtn.Click += (_, _) => Close();
        Controls.Add(backBtn);
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
