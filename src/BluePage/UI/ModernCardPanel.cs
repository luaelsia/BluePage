using System.Drawing.Drawing2D;

namespace Microsoft365OfficeWebLauncher.UI;

internal sealed class ModernCardPanel : Panel
{
    public int CornerRadius { get; set; } = 12;

    public ModernCardPanel()
    {
        DoubleBuffered = true;
        Resize += (_, _) => UpdateRegion();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        using var path = CreateRoundedPath(ClientRectangle, CornerRadius);
        using var pen = new Pen(AppTheme.Current.Border);
        e.Graphics.DrawPath(pen, path);
    }

    private void UpdateRegion()
    {
        if (Width <= 0 || Height <= 0)
        {
            return;
        }

        using var path = CreateRoundedPath(ClientRectangle, CornerRadius);
        Region?.Dispose();
        Region = new Region(path);
        Invalidate();
    }

    private static GraphicsPath CreateRoundedPath(Rectangle bounds, int radius)
    {
        var path = new GraphicsPath();
        var rect = Rectangle.Inflate(bounds, -1, -1);
        var diameter = Math.Min(radius * 2, Math.Min(rect.Width, rect.Height));
        var arc = new Rectangle(rect.Location, new Size(diameter, diameter));

        path.AddArc(arc, 180, 90);
        arc.X = rect.Right - diameter;
        path.AddArc(arc, 270, 90);
        arc.Y = rect.Bottom - diameter;
        path.AddArc(arc, 0, 90);
        arc.X = rect.Left;
        path.AddArc(arc, 90, 90);
        path.CloseFigure();
        return path;
    }
}
