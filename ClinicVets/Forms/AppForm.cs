using System.Windows.Forms;

namespace ClinicVets.Forms;

/// <summary>
/// Base form for every window in the app.
///
/// It eliminates the "whole window flashes / looks like it is refreshing"
/// effect that happens when a WinForms form is maximized or resized while it
/// contains many child controls. Two things make that happen:
///   1. The form paints to an off-screen buffer (double buffering).
///   2. WS_EX_COMPOSITED tells Windows to buffer the painting of ALL child
///      controls together, so resizing repaints the window in one shot
///      instead of control-by-control (which is what looks like flicker).
/// </summary>
public class AppForm : Form
{
    public AppForm()
    {
        // Paint the form itself through a back buffer.
        DoubleBuffered = true;
        SetStyle(
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.AllPaintingInWmPaint,
            true);
    }

    /// <summary>
    /// WS_EX_COMPOSITED removes resize flicker, but a few controls
    /// (notably CheckedListBox / ListBox) render their text incorrectly under
    /// it — the text looks like it is "smearing" or refreshing. Forms that host
    /// such controls override this to false and rely on plain double buffering.
    /// </summary>
    protected virtual bool UseCompositedPainting => true;

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            // WS_EX_COMPOSITED — buffer the whole window (children included).
            if (UseCompositedPainting)
                cp.ExStyle |= 0x02000000;
            return cp;
        }
    }
}
