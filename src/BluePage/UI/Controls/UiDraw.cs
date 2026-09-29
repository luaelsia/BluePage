using System.Drawing.Drawing2D;

namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>직접 그리는 컨트롤들이 함께 쓰는 DPI 배율, 둥근 사각형, 아이콘 글꼴.</summary>
internal static class UiDraw
{
    private static readonly Lazy<float> ScaleFactor = new(() =>
    {
        using var graphics = Graphics.FromHwnd(IntPtr.Zero);
        return graphics.DpiX / 96F;
    });

    private static readonly Lazy<string> IconFontFamily = new(() =>
    {
        // Windows 11은 Segoe Fluent Icons, Windows 10은 Segoe MDL2 Assets. 두 글꼴은 같은 코드 포인트를 쓴다.
        using var fonts = new System.Drawing.Text.InstalledFontCollection();
        return fonts.Families.Any(family => family.Name == "Segoe Fluent Icons")
            ? "Segoe Fluent Icons"
            : "Segoe MDL2 Assets";
    });

    /// <summary>96 DPI 기준 픽셀 값을 현재 DPI로 바꾼다. 글꼴은 이미 DPI에 맞춰 커지므로 고정 크기에만 쓴다.</summary>
    public static int S(int pixels) => (int)Math.Round(pixels * ScaleFactor.Value);

    public static Padding S(int left, int top, int right, int bottom) => new(S(left), S(top), S(right), S(bottom));

    public static Font IconFont(float size) => new(IconFontFamily.Value, size, FontStyle.Regular, GraphicsUnit.Point);

    public static GraphicsPath RoundedRect(RectangleF bounds, float radius)
    {
        var path = new GraphicsPath();
        var diameter = Math.Min(radius * 2, Math.Min(bounds.Width, bounds.Height));
        if (diameter <= 0)
        {
            path.AddRectangle(bounds);
            return path;
        }
        var arc = new RectangleF(bounds.Location, new SizeF(diameter, diameter));
        path.AddArc(arc, 180, 90);
        arc.X = bounds.Right - diameter;
        path.AddArc(arc, 270, 90);
        arc.Y = bounds.Bottom - diameter;
        path.AddArc(arc, 0, 90);
        arc.X = bounds.Left;
        path.AddArc(arc, 90, 90);
        path.CloseFigure();
        return path;
    }

    /// <summary>부모의 배경색. 카드 안에 있으면 카드 색, 아니면 부모 BackColor.</summary>
    public static Color SurfaceBehind(Control control) =>
        control.Parent is null ? AppTheme.Current.Background : control.Parent.BackColor;

    public static Color Blend(Color first, Color second, float amount) => Color.FromArgb(
        (int)(first.R + (second.R - first.R) * amount),
        (int)(first.G + (second.G - first.G) * amount),
        (int)(first.B + (second.B - first.B) * amount));
}

/// <summary>Segoe Fluent Icons / MDL2 Assets 코드 포인트.</summary>
internal static class Glyphs
{
    public const string Home = "";
    public const string Document = "";
    public const string Settings = "";
    public const string Sync = "";
    public const string Link = "";
    public const string ChevronDown = "";
    public const string CheckMark = "";
    public const string Info = "";
}
