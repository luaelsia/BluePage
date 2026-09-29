namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>알약 모양 트랙 안에서 하나를 고르는 선택 버튼 묶음(예: 시스템 / 라이트 / 다크).</summary>
internal sealed class SegmentedControl : Control
{
    private readonly string[] _items;
    private readonly List<RectangleF> _segments = new();
    private int _selectedIndex = -1;
    private int _hoverIndex = -1;

    public event EventHandler? SelectedIndexChanged;

    public SegmentedControl(params string[] items)
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
        _items = items;
        TabStop = true;
        Cursor = Cursors.Hand;
        Margin = new Padding(0);
        RecalculateSegments();
    }

    public int SelectedIndex
    {
        get => _selectedIndex;
        set
        {
            if (value == _selectedIndex || value < -1 || value >= _items.Length)
            {
                return;
            }
            _selectedIndex = value;
            Invalidate();
            SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    protected override void OnFontChanged(EventArgs e)
    {
        base.OnFontChanged(e);
        RecalculateSegments();
    }

    private void RecalculateSegments()
    {
        _segments.Clear();
        var pad = UiDraw.S(3);
        var x = (float)pad;
        var height = TextRenderer.MeasureText("가", Font).Height + UiDraw.S(12);
        foreach (var item in _items)
        {
            var width = TextRenderer.MeasureText(item, Font, Size.Empty, TextFormatFlags.NoPadding).Width + UiDraw.S(26);
            _segments.Add(new RectangleF(x, pad, width, height));
            x += width;
        }
        Size = new Size((int)x + pad, height + pad * 2);
        Invalidate();
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        var index = _segments.FindIndex(segment => segment.Contains(e.Location));
        if (index != _hoverIndex)
        {
            _hoverIndex = index;
            Invalidate();
        }
        base.OnMouseMove(e);
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        _hoverIndex = -1;
        Invalidate();
        base.OnMouseLeave(e);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        Focus();
        var index = _segments.FindIndex(segment => segment.Contains(e.Location));
        if (index >= 0)
        {
            SelectedIndex = index;
        }
        base.OnMouseDown(e);
    }

    protected override bool IsInputKey(Keys keyData) => keyData is Keys.Left or Keys.Right || base.IsInputKey(keyData);

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Left && _selectedIndex > 0) SelectedIndex = _selectedIndex - 1;
        if (e.KeyCode == Keys.Right && _selectedIndex < _items.Length - 1) SelectedIndex = _selectedIndex + 1;
        base.OnKeyDown(e);
    }

    protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
    protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }

    protected override void OnPaint(PaintEventArgs e)
    {
        var theme = AppTheme.Current;
        var g = e.Graphics;
        var surface = UiDraw.SurfaceBehind(this);
        g.Clear(surface);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        var trackColor = UiDraw.Blend(surface, theme.TextPrimary, AppTheme.IsDark ? 0.08F : 0.05F);
        using (var track = UiDraw.RoundedRect(new RectangleF(0, 0, Width - 1, Height - 1), (Height - 1) / 2F))
        using (var fill = new SolidBrush(trackColor))
        {
            g.FillPath(fill, track);
            if (Focused)
            {
                using var focus = new Pen(theme.Accent, 1.2F);
                g.DrawPath(focus, track);
            }
        }

        for (var i = 0; i < _items.Length; i++)
        {
            var segment = _segments[i];
            var selected = i == _selectedIndex;
            if (selected || i == _hoverIndex)
            {
                using var path = UiDraw.RoundedRect(segment, segment.Height / 2);
                using var fill = new SolidBrush(selected ? theme.CardBackground : UiDraw.Blend(trackColor, theme.TextPrimary, 0.05F));
                g.FillPath(fill, path);
                if (selected)
                {
                    using var border = new Pen(theme.Border);
                    g.DrawPath(border, path);
                }
            }
            TextRenderer.DrawText(g, _items[i], Font, Rectangle.Round(segment),
                selected ? theme.TextPrimary : theme.TextSecondary,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
        }
    }
}
