namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>둥근 모서리 카드. 안쪽 컨트롤은 카드 배경색을 그대로 물려받는다(ThemeApplier 참고).</summary>
internal sealed class CardPanel : Panel
{
    public int CornerRadius { get; set; } = UiDraw.S(14);

    public CardPanel()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
        Padding = UiDraw.S(18, 16, 18, 16);
        BackColor = AppTheme.Current.CardBackground;
    }

    protected override void OnPaintBackground(PaintEventArgs e)
    {
        var theme = AppTheme.Current;
        e.Graphics.Clear(UiDraw.SurfaceBehind(this));
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        var bounds = new RectangleF(0.5F, 0.5F, Width - 1.5F, Height - 1.5F);
        using var path = UiDraw.RoundedRect(bounds, CornerRadius);
        using var fill = new SolidBrush(theme.CardBackground);
        e.Graphics.FillPath(fill, path);
        using var border = new Pen(theme.Border);
        e.Graphics.DrawPath(border, path);
    }
}
