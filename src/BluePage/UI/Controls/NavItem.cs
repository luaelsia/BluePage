namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>왼쪽 사이드 메뉴의 항목 하나. 선택되면 연한 강조색 배경과 강조색 글자로 그린다.</summary>
internal sealed class NavItem : Control
{
    private readonly string _glyph;
    private bool _selected;
    private bool _hovered;

    public NavItem(string glyph, string text)
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
        _glyph = glyph;
        Text = text;
        TabStop = true;
        Cursor = Cursors.Hand;
        Height = UiDraw.S(40);
        Margin = new Padding(0, 0, 0, UiDraw.S(4));
        AccessibleRole = AccessibleRole.PageTab;
        AccessibleName = text;
    }

    public bool Selected
    {
        get => _selected;
        set { _selected = value; Invalidate(); }
    }

    protected override void OnMouseEnter(EventArgs e) { _hovered = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { _hovered = false; Invalidate(); base.OnMouseLeave(e); }
    protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
    protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (e.KeyCode is Keys.Enter or Keys.Space)
        {
            OnClick(EventArgs.Empty);
        }
        base.OnKeyDown(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var theme = AppTheme.Current;
        var g = e.Graphics;
        var surface = UiDraw.SurfaceBehind(this);
        g.Clear(surface);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        if (_selected || _hovered || (Focused && ShowFocusCues))
        {
            var back = _selected ? theme.AccentSoft : UiDraw.Blend(surface, theme.TextPrimary, 0.05F);
            using var path = UiDraw.RoundedRect(new RectangleF(0, 0, Width - 1, Height - 1), UiDraw.S(10));
            using var fill = new SolidBrush(back);
            g.FillPath(fill, path);
        }

        var color = _selected ? theme.AccentSoftText : theme.TextSecondary;
        using var iconFont = UiDraw.IconFont(11F);
        TextRenderer.DrawText(g, _glyph, iconFont, new Rectangle(UiDraw.S(12), 0, UiDraw.S(22), Height), color,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);

        using var font = _selected ? new Font(Font, FontStyle.Bold) : null;
        TextRenderer.DrawText(g, Text, font ?? Font, new Rectangle(UiDraw.S(44), 0, Width - UiDraw.S(48), Height),
            _selected ? theme.AccentSoftText : theme.TextPrimary,
            TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis);
    }
}
