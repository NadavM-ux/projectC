using System.Windows.Forms;
using ClinicVets.Data;
using ClinicVets.Forms;

namespace ClinicVets;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();

        var dataPath = Path.Combine(AppContext.BaseDirectory, "ClinicVetsData.xlsx");
        var store = new ExcelDataStore(dataPath);
        store.Load();

        while (true)
        {
            using var login = new LoginForm(store);
            var result = login.ShowDialog();
            if (result != DialogResult.OK || login.AuthenticatedEmployee == null)
            {
                return;
            }

            using var main = new MainForm(store, login.AuthenticatedEmployee);
            Application.Run(main);
        }
    }
}
