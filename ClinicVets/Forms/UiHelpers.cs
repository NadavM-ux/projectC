using System.Drawing;
using System.Windows.Forms;

namespace ClinicVets.Forms;

internal static class UiHelpers
{
    public static Button MakeButton(string text, Point location, Color back, Color? fore = null, Size? size = null)
    {
        var b = new Button
        {
            Text = text,
            Location = location,
            Size = size ?? new Size(90, 30),
            BackColor = back,
            ForeColor = fore ?? Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9F, FontStyle.Bold),
        };
        b.FlatAppearance.BorderSize = 0;
        return b;
    }
}
