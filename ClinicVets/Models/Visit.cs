namespace ClinicVets.Models;

public class Visit
{
    public string VisitId { get; set; } = "";
    public string AnimalChipNumber { get; set; } = "";
    public string VetUsername { get; set; } = "";
    public DateTime VisitDateTime { get; set; }
    public string Reason { get; set; } = "";
    public string Diagnosis { get; set; } = "";
    public List<string> MedicationsGiven { get; set; } = new();
    public decimal BasePrice { get; set; }
    public decimal TotalPrice { get; set; }
}
