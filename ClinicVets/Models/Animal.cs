namespace ClinicVets.Models;

public class Animal
{
    public string ChipNumber { get; set; } = "";
    public string Name { get; set; } = "";
    public string Species { get; set; } = "";
    public double Weight { get; set; }
    public DateTime DateOfBirth { get; set; }
    public string OwnerNationalId { get; set; } = "";
    public DateTime? LastVaccinationDate { get; set; }
}
