using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ClinicVets.Forms;

internal static class UiHelpers
{
    // Turns on double buffering for a control that does not expose it publicly
    // (e.g. DataGridView). This stops the grid from flickering when it is
    // redrawn during a window resize or when its data is refreshed.
    public static void EnableDoubleBuffering(Control control)
    {
        var prop = typeof(Control).GetProperty(
            "DoubleBuffered",
            BindingFlags.Instance | BindingFlags.NonPublic);
        prop?.SetValue(control, true, null);
    }

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

    // Lets an absolute-positioned form grow to full screen without breaking its layout.
    // The control bounds present when the form first loads are treated as the design layout;
    // on every resize the whole layout is scaled by a single uniform factor (so nothing is
    // stretched or distorted) and centered in the client area.
    public static void EnableProportionalScaling(Form form)
    {
        var design = new Dictionary<Control, Rectangle>();
        // Remember the original font of every control so we can scale the text
        // by the same factor as the boxes — otherwise the boxes grow but the
        // text stays small.
        var designFont = new Dictionary<Control, Font>();
        var designClient = Size.Empty;
        var captured = false;

        void Capture(Control parent)
        {
            foreach (Control c in parent.Controls)
            {
                design[c] = c.Bounds;
                designFont[c] = c.Font;
                Capture(c);
            }
        }

        void ApplyTo(Control parent, float scale, Point offset)
        {
            foreach (Control c in parent.Controls)
            {
                // Scale the text to match the new size of the control.
                if (designFont.TryGetValue(c, out var of))
                {
                    float newSize = Math.Max(1f, of.Size * scale);
                    if (Math.Abs(c.Font.Size - newSize) > 0.1f)
                        c.Font = new Font(of.FontFamily, newSize, of.Style);
                }

                // Docked controls are positioned by the layout engine, not by us.
                if (c.Dock != DockStyle.None)
                {
                    ApplyTo(c, scale, Point.Empty);
                    continue;
                }
                if (!design.TryGetValue(c, out var b)) continue;
                c.Bounds = new Rectangle(
                    (int)Math.Round(b.X * scale) + offset.X,
                    (int)Math.Round(b.Y * scale) + offset.Y,
                    (int)Math.Round(b.Width * scale),
                    (int)Math.Round(b.Height * scale));
                ApplyTo(c, scale, Point.Empty);
            }
        }

        void Apply()
        {
            if (!captured || designClient.Width == 0 || designClient.Height == 0) return;
            float scale = Math.Min(
                (float)form.ClientSize.Width / designClient.Width,
                (float)form.ClientSize.Height / designClient.Height);
            int offX = (int)Math.Round((form.ClientSize.Width - designClient.Width * scale) / 2f);
            int offY = (int)Math.Round((form.ClientSize.Height - designClient.Height * scale) / 2f);
            form.SuspendLayout();
            ApplyTo(form, scale, new Point(offX, offY));
            form.ResumeLayout(true);
        }

        form.Load += (_, _) =>
        {
            if (captured) return;
            designClient = form.ClientSize;
            Capture(form);
            captured = true;
            Apply();
        };
        form.Resize += (_, _) => Apply();
    }
}
