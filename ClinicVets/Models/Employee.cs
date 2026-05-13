namespace ClinicVets.Models;

public enum Role
{
    Veterinarian,
    Secretary
}

public class Employee
{
    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
    public string EmployeeNumber { get; set; } = "";
    public string FullName { get; set; } = "";
    public string Email { get; set; } = "";
    public string NationalId { get; set; } = "";
    public Role Role { get; set; }
}
