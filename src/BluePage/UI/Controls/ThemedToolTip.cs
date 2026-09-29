namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>테마 색으로 그리는 툴팁. 설정 항목 설명을 화면에 늘어놓지 않고 마우스를 올렸을 때 보여 준다.</summary>
internal sealed class ThemedToolTip : ToolTip
{
    private const TextFormatFlags Flags = TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix | TextFormatFlags.Left;
    private readonly Font _font = new("Segoe UI Variable Text", 9F);

    public ThemedToolTip()
    {
        OwnerDraw = true;
        InitialDelay = 250;
        ReshowDelay = 100;
        AutoPopDelay = 20000;
        ShowAlways = true;
        Popup += OnPopup;
        Draw += OnDraw;
    }

    private void OnPopup(object? sender, PopupEventArgs e)
    {
        var text = e.AssociatedControl is null ? string.Empty : GetToolTip(e.AssociatedControl);
        var size = TextRenderer.MeasureText(text, _font, new Size(UiDraw.S(300), 0), Flags);
        e.ToolTipSize = new Size(size.Width + UiDraw.S(24), size.Height + UiDraw.S(16));
    }

    private void OnDraw(object? sender, DrawToolTipEventArgs e)
    {
        var theme = AppTheme.Current;
        using (var back = new SolidBrush(theme.CardBackground))
        {
            e.Graphics.FillRectangle(back, e.Bounds);
        }
        using (var border = new Pen(theme.Border))
        {
            e.Graphics.DrawRectangle(border, 0, 0, e.Bounds.Width - 1, e.Bounds.Height - 1);
        }
        var textBounds = Rectangle.Inflate(e.Bounds, -UiDraw.S(12), -UiDraw.S(8));
        TextRenderer.DrawText(e.Graphics, e.ToolTipText, _font, textBounds, theme.TextPrimary, Flags);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _font.Dispose();
        }
        base.Dispose(disposing);
    }
}
