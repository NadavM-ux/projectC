# ClinicVets

Veterinary clinic management system. C# Windows Forms (.NET 8) with Excel storage via ClosedXML.

## Run

In Visual Studio (Windows):
1. Open `ClinicVets.sln`
2. F5 to run

From command line on Windows:
```
dotnet run --project ClinicVets
```

The data file `ClinicVetsData.xlsx` is created next to the executable on first run, with seeded medications.

## First use

1. On first launch, the Employees sheet is empty — click **Register** to create the first employee.
2. Log in with your credentials.
3. Buttons on the main hub are gated by role:
   - **Customers** — Secretary only
   - **Animals** — All staff
   - **Visits / Medications** — Veterinarian only

## Validation rules (per spec)

| Field | Rule |
|---|---|
| Username | 6–8 chars, max 2 digits, rest English letters |
| Password | 8–10 chars, ≥1 letter, ≥1 digit, ≥1 of `! # $` |
| Employee number | exactly 4 digits |
| National ID | 9 digits + valid Israeli checksum |
| Email | must contain `@` and a valid TLD |
| Animal name | letters only |
| Weight | 0.1–100 kg |
| Date of birth | not in the future, not before 2000 |

## Project layout

```
ClinicVets/
  Models/        Employee, Customer, Animal, Visit, Medication
  Validation/    Validators (rule checks)
  Data/          ExcelDataStore (ClosedXML CRUD)
  Forms/         Login, Register, Main, Customers, Animals, Visits, Medications
  Program.cs     Entry point
```
