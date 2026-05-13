using System.Text.RegularExpressions;

namespace ClinicVets.Validation;

public static class Validators
{
    public static bool IsValidUsername(string username, out string error)
    {
        error = "";
        if (string.IsNullOrWhiteSpace(username))
        {
            error = "Username is required.";
            return false;
        }
        if (username.Length < 6 || username.Length > 8)
        {
            error = "Username must be 6-8 characters.";
            return false;
        }
        int digitCount = 0;
        foreach (char c in username)
        {
            if (char.IsDigit(c))
            {
                digitCount++;
            }
            else if (!IsEnglishLetter(c))
            {
                error = "Username may contain only English letters and digits.";
                return false;
            }
        }
        if (digitCount > 2)
        {
            error = "Username may contain at most 2 digits.";
            return false;
        }
        return true;
    }

    public static bool IsValidPassword(string password, out string error)
    {
        error = "";
        if (string.IsNullOrEmpty(password))
        {
            error = "Password is required.";
            return false;
        }
        if (password.Length < 8 || password.Length > 10)
        {
            error = "Password must be 8-10 characters.";
            return false;
        }
        bool hasLetter = false;
        bool hasDigit = false;
        bool hasSpecial = false;
        foreach (char c in password)
        {
            if (char.IsLetter(c)) hasLetter = true;
            else if (char.IsDigit(c)) hasDigit = true;
            else if (c == '!' || c == '#' || c == '$') hasSpecial = true;
        }
        if (!hasLetter || !hasDigit || !hasSpecial)
        {
            error = "Password must contain at least one letter, one digit and one special character (!, #, $).";
            return false;
        }
        return true;
    }

    public static bool IsValidEmployeeNumber(string number, out string error)
    {
        error = "";
        if (number == null || number.Length != 4 || !number.All(char.IsDigit))
        {
            error = "Employee number must be exactly 4 digits.";
            return false;
        }
        return true;
    }

    public static bool IsValidEmail(string email, out string error)
    {
        error = "";
        if (string.IsNullOrWhiteSpace(email))
        {
            error = "Email is required.";
            return false;
        }
        var pattern = @"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$";
        if (!Regex.IsMatch(email, pattern))
        {
            error = "Invalid email format.";
            return false;
        }
        return true;
    }

    public static bool IsValidNationalId(string id, out string error)
    {
        error = "";
        if (string.IsNullOrWhiteSpace(id) || id.Length != 9 || !id.All(char.IsDigit))
        {
            error = "National ID must be exactly 9 digits.";
            return false;
        }
        if (!IsraeliIdChecksumValid(id))
        {
            error = "National ID checksum is invalid.";
            return false;
        }
        return true;
    }

    public static bool IsValidPhone(string phone, out string error)
    {
        error = "";
        if (string.IsNullOrWhiteSpace(phone))
        {
            error = "Phone is required.";
            return false;
        }
        var digits = new string(phone.Where(char.IsDigit).ToArray());
        if (digits.Length < 9 || digits.Length > 10)
        {
            error = "Phone must contain 9-10 digits.";
            return false;
        }
        return true;
    }

    public static bool IsLettersOnly(string value, out string error, string fieldName = "Field")
    {
        error = "";
        if (string.IsNullOrWhiteSpace(value))
        {
            error = $"{fieldName} is required.";
            return false;
        }
        foreach (char c in value)
        {
            if (!char.IsLetter(c) && c != ' ' && c != '-' && c != '\'')
            {
                error = $"{fieldName} may contain only letters.";
                return false;
            }
        }
        return true;
    }

    public static bool IsValidWeight(double weight, out string error)
    {
        error = "";
        if (weight < 0.1 || weight > 100)
        {
            error = "Weight must be between 0.1 and 100 kg.";
            return false;
        }
        return true;
    }

    public static bool IsValidBirthDate(DateTime dob, out string error)
    {
        error = "";
        if (dob > DateTime.Now)
        {
            error = "Date of birth cannot be in the future.";
            return false;
        }
        if (dob.Year < 2000)
        {
            error = "Date of birth cannot be before the year 2000.";
            return false;
        }
        return true;
    }

    private static bool IsEnglishLetter(char c)
    {
        return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
    }

    private static bool IsraeliIdChecksumValid(string id)
    {
        int sum = 0;
        for (int i = 0; i < 9; i++)
        {
            int digit = id[i] - '0';
            int multiplier = (i % 2) + 1;
            int product = digit * multiplier;
            if (product > 9) product -= 9;
            sum += product;
        }
        return sum % 10 == 0;
    }
}
