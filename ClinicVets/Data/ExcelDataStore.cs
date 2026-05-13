using System.Globalization;
using ClinicVets.Models;
using ClosedXML.Excel;

namespace ClinicVets.Data;

public class ExcelDataStore
{
    private readonly string _filePath;

    public List<Employee> Employees { get; private set; } = new();
    public List<Customer> Customers { get; private set; } = new();
    public List<Animal> Animals { get; private set; } = new();
    public List<Visit> Visits { get; private set; } = new();
    public List<Medication> Medications { get; private set; } = new();
    public List<string> Species { get; private set; } = new();

    private static readonly string[] DefaultSpecies = { "Dog", "Cat", "Reptile", "Bird" };

    public ExcelDataStore(string filePath)
    {
        _filePath = filePath;
    }

    public void Load()
    {
        if (!File.Exists(_filePath))
        {
            SeedDefaults();
            Save();
            return;
        }

        using var wb = new XLWorkbook(_filePath);
        Employees = ReadEmployees(wb);
        Customers = ReadCustomers(wb);
        Animals = ReadAnimals(wb);
        Visits = ReadVisits(wb);
        Medications = ReadMedications(wb);
        Species = ReadSpecies(wb);

        bool dirty = false;
        if (Medications.Count == 0) { SeedMedications(); dirty = true; }
        if (Species.Count == 0) { SeedSpecies(); dirty = true; }
        if (dirty) Save();
    }

    public void Save()
    {
        using var wb = new XLWorkbook();
        WriteEmployees(wb);
        WriteCustomers(wb);
        WriteAnimals(wb);
        WriteVisits(wb);
        WriteMedications(wb);
        WriteSpecies(wb);
        wb.SaveAs(_filePath);
    }

    public bool TryAddAnimal(Animal animal, out string error)
    {
        error = "";
        if (Animals.Any(a => a.ChipNumber == animal.ChipNumber))
        {
            error = "Chip number already exists.";
            return false;
        }
        Animals.Add(animal);
        Save();
        return true;
    }

    public bool TryAddSpecies(string name, out string error)
    {
        error = "";
        var trimmed = name.Trim();
        if (string.IsNullOrEmpty(trimmed))
        {
            error = "Species name is required.";
            return false;
        }
        if (Species.Any(s => string.Equals(s, trimmed, StringComparison.OrdinalIgnoreCase)))
        {
            error = "Species already exists.";
            return false;
        }
        Species.Add(trimmed);
        Save();
        return true;
    }

    public bool RemoveSpecies(string name)
    {
        var idx = Species.FindIndex(s => string.Equals(s, name, StringComparison.OrdinalIgnoreCase));
        if (idx < 0) return false;
        Species.RemoveAt(idx);
        Save();
        return true;
    }

    private void SeedDefaults()
    {
        SeedMedications();
        SeedSpecies();
    }

    private void SeedSpecies()
    {
        Species = DefaultSpecies.ToList();
    }

    private void SeedMedications()
    {
        Medications = new List<Medication>
        {
            new() { Name = "Antibiotics", Price = 80m, StockQuantity = 50 },
            new() { Name = "Painkiller", Price = 45m, StockQuantity = 80 },
            new() { Name = "Vaccine - Annual", Price = 120m, StockQuantity = 30 },
            new() { Name = "Anti-parasitic", Price = 60m, StockQuantity = 40 },
        };
    }

    // ---------- Employees ----------
    private static List<Employee> ReadEmployees(XLWorkbook wb)
    {
        var list = new List<Employee>();
        if (!wb.TryGetWorksheet("Employees", out var ws)) return list;
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            list.Add(new Employee
            {
                Username = row.Cell(1).GetString(),
                Password = row.Cell(2).GetString(),
                EmployeeNumber = row.Cell(3).GetString(),
                FullName = row.Cell(4).GetString(),
                Email = row.Cell(5).GetString(),
                NationalId = row.Cell(6).GetString(),
                Role = Enum.TryParse<Role>(row.Cell(7).GetString(), out var r) ? r : Role.Secretary,
            });
        }
        return list;
    }

    private void WriteEmployees(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Employees");
        WriteHeader(ws, "Username", "Password", "EmployeeNumber", "FullName", "Email", "NationalId", "Role");
        for (int i = 0; i < Employees.Count; i++)
        {
            var e = Employees[i];
            var row = ws.Row(i + 2);
            row.Cell(1).Value = e.Username;
            row.Cell(2).Value = e.Password;
            row.Cell(3).Value = e.EmployeeNumber;
            row.Cell(4).Value = e.FullName;
            row.Cell(5).Value = e.Email;
            row.Cell(6).Value = e.NationalId;
            row.Cell(7).Value = e.Role.ToString();
        }
    }

    // ---------- Customers ----------
    private static List<Customer> ReadCustomers(XLWorkbook wb)
    {
        var list = new List<Customer>();
        if (!wb.TryGetWorksheet("Customers", out var ws)) return list;
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            list.Add(new Customer
            {
                FullName = row.Cell(1).GetString(),
                NationalId = row.Cell(2).GetString(),
                Phone = row.Cell(3).GetString(),
                Email = row.Cell(4).GetString(),
            });
        }
        return list;
    }

    private void WriteCustomers(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Customers");
        WriteHeader(ws, "FullName", "NationalId", "Phone", "Email");
        for (int i = 0; i < Customers.Count; i++)
        {
            var c = Customers[i];
            var row = ws.Row(i + 2);
            row.Cell(1).Value = c.FullName;
            row.Cell(2).Value = c.NationalId;
            row.Cell(3).Value = c.Phone;
            row.Cell(4).Value = c.Email;
        }
    }

    // ---------- Animals ----------
    private static List<Animal> ReadAnimals(XLWorkbook wb)
    {
        var list = new List<Animal>();
        if (!wb.TryGetWorksheet("Animals", out var ws)) return list;
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            list.Add(new Animal
            {
                ChipNumber = row.Cell(1).GetString(),
                Name = row.Cell(2).GetString(),
                Species = row.Cell(3).GetString(),
                Weight = double.TryParse(row.Cell(4).GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var w) ? w : 0,
                DateOfBirth = DateTime.TryParse(row.Cell(5).GetString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : DateTime.MinValue,
                OwnerNationalId = row.Cell(6).GetString(),
                LastVaccinationDate = DateTime.TryParse(row.Cell(7).GetString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var v) ? v : null,
            });
        }
        return list;
    }

    private void WriteAnimals(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Animals");
        WriteHeader(ws, "ChipNumber", "Name", "Species", "Weight", "DateOfBirth", "OwnerNationalId", "LastVaccinationDate");
        for (int i = 0; i < Animals.Count; i++)
        {
            var a = Animals[i];
            var row = ws.Row(i + 2);
            row.Cell(1).Value = a.ChipNumber;
            row.Cell(2).Value = a.Name;
            row.Cell(3).Value = a.Species;
            row.Cell(4).Value = a.Weight.ToString(CultureInfo.InvariantCulture);
            row.Cell(5).Value = a.DateOfBirth.ToString("yyyy-MM-dd");
            row.Cell(6).Value = a.OwnerNationalId;
            row.Cell(7).Value = a.LastVaccinationDate?.ToString("yyyy-MM-dd") ?? "";
        }
    }

    // ---------- Visits ----------
    private static List<Visit> ReadVisits(XLWorkbook wb)
    {
        var list = new List<Visit>();
        if (!wb.TryGetWorksheet("Visits", out var ws)) return list;
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            var medsCsv = row.Cell(7).GetString();
            list.Add(new Visit
            {
                VisitId = row.Cell(1).GetString(),
                AnimalChipNumber = row.Cell(2).GetString(),
                VetUsername = row.Cell(3).GetString(),
                VisitDateTime = DateTime.TryParse(row.Cell(4).GetString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt) ? dt : DateTime.MinValue,
                Reason = row.Cell(5).GetString(),
                Diagnosis = row.Cell(6).GetString(),
                MedicationsGiven = string.IsNullOrWhiteSpace(medsCsv)
                    ? new List<string>()
                    : medsCsv.Split('|', StringSplitOptions.RemoveEmptyEntries).ToList(),
                BasePrice = decimal.TryParse(row.Cell(8).GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var bp) ? bp : 0m,
                TotalPrice = decimal.TryParse(row.Cell(9).GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var tp) ? tp : 0m,
            });
        }
        return list;
    }

    private void WriteVisits(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Visits");
        WriteHeader(ws, "VisitId", "AnimalChipNumber", "VetUsername", "VisitDateTime", "Reason", "Diagnosis", "Medications", "BasePrice", "TotalPrice");
        for (int i = 0; i < Visits.Count; i++)
        {
            var v = Visits[i];
            var row = ws.Row(i + 2);
            row.Cell(1).Value = v.VisitId;
            row.Cell(2).Value = v.AnimalChipNumber;
            row.Cell(3).Value = v.VetUsername;
            row.Cell(4).Value = v.VisitDateTime.ToString("yyyy-MM-dd HH:mm");
            row.Cell(5).Value = v.Reason;
            row.Cell(6).Value = v.Diagnosis;
            row.Cell(7).Value = string.Join("|", v.MedicationsGiven);
            row.Cell(8).Value = v.BasePrice.ToString(CultureInfo.InvariantCulture);
            row.Cell(9).Value = v.TotalPrice.ToString(CultureInfo.InvariantCulture);
        }
    }

    // ---------- Medications ----------
    private static List<Medication> ReadMedications(XLWorkbook wb)
    {
        var list = new List<Medication>();
        if (!wb.TryGetWorksheet("Medications", out var ws)) return list;
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            list.Add(new Medication
            {
                Name = row.Cell(1).GetString(),
                Price = decimal.TryParse(row.Cell(2).GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var p) ? p : 0m,
                StockQuantity = int.TryParse(row.Cell(3).GetString(), out var s) ? s : 0,
            });
        }
        return list;
    }

    private void WriteMedications(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Medications");
        WriteHeader(ws, "Name", "Price", "StockQuantity");
        for (int i = 0; i < Medications.Count; i++)
        {
            var m = Medications[i];
            var row = ws.Row(i + 2);
            row.Cell(1).Value = m.Name;
            row.Cell(2).Value = m.Price.ToString(CultureInfo.InvariantCulture);
            row.Cell(3).Value = m.StockQuantity;
        }
    }

    // ---------- Species ----------
    private static List<string> ReadSpecies(XLWorkbook wb)
    {
        var list = new List<string>();
        if (!wb.TryGetWorksheet("Species", out var ws)) return list;
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            var name = row.Cell(1).GetString().Trim();
            if (!string.IsNullOrEmpty(name)) list.Add(name);
        }
        return list;
    }

    private void WriteSpecies(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Species");
        WriteHeader(ws, "Name");
        for (int i = 0; i < Species.Count; i++)
        {
            ws.Row(i + 2).Cell(1).Value = Species[i];
        }
    }

    private static void WriteHeader(IXLWorksheet ws, params string[] columns)
    {
        for (int i = 0; i < columns.Length; i++)
        {
            var cell = ws.Cell(1, i + 1);
            cell.Value = columns[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.LightGray;
        }
    }
}
