namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>
/// 콤보박스 대신 쓰는 드롭다운. 누르면 테마 색 메뉴로 항목을 보여 주고, 고른 항목에 체크 표시를 한다.
/// ComboBox와 같은 SelectedIndex / SelectedIndexChanged를 제공한다.
/// </summary>
internal sealed class SelectBox : Control
{
    /// <summary>기본 폭. 선택 상자 대신 글자를 둘 때도 이 폭에 맞춰 좌우 위치를 맞춘다.</summary>
    public static int DefaultWidth => UiDraw.S(220);

    private readonly string[] _items;
    private int _selectedIndex = -1;
    private bool _hovered;

    public event EventHandler? SelectedIndexChanged;

    public SelectBox(params string[] items)
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
        _items = items;
        TabStop = true;
        Cursor = Cursors.Hand;
        Margin = new Padding(0);
        Width = DefaultWidth;
        Height = TextRenderer.MeasureText("가", Font).Height + UiDraw.S(16);
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
        Height = TextRenderer.MeasureText("가", Font).Height + UiDraw.S(16);
    }

    protected override void OnMouseEnter(EventArgs e) { _hovered = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { _hovered = false; Invalidate(); base.OnMouseLeave(e); }
    protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
    protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }
    protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (e.Button == MouseButtons.Left)
        {
            Focus();
            ShowMenu();
        }
    }

    protected override bool IsInputKey(Keys keyData) => keyData is Keys.Up or Keys.Down || base.IsInputKey(keyData);

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (e.KeyCode is Keys.Space or Keys.Enter || (e.Alt && e.KeyCode == Keys.Down)) ShowMenu();
        else if (e.KeyCode == Keys.Down && _selectedIndex < _items.Length - 1) SelectedIndex = _selectedIndex + 1;
        else if (e.KeyCode == Keys.Up && _selectedIndex > 0) SelectedIndex = _selectedIndex - 1;
        base.OnKeyDown(e);
    }

    private void ShowMenu()
    {
        var menu = new ContextMenuStrip { ShowImageMargin = false, ShowCheckMargin = true, Font = Font };
        for (var i = 0; i < _items.Length; i++)
        {
            var index = i;
            var item = new ToolStripMenuItem(_items[i]) { Checked = i == _selectedIndex, Padding = new Padding(0, UiDraw.S(3), 0, UiDraw.S(3)) };
            item.Click += (_, _) => SelectedIndex = index;
            menu.Items.Add(item);
        }
        ThemedMenu.Apply(menu, centerText: true);
        menu.MinimumSize = new Size(Width, 0);
        menu.Closed += (_, _) => BeginInvoke(menu.Dispose);
        menu.Show(this, new Point(0, Height + UiDraw.S(2)));
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var theme = AppTheme.Current;
        var g = e.Graphics;
        g.Clear(UiDraw.SurfaceBehind(this));
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        // 테두리 없이 바탕보다 살짝 밝은(라이트에서는 살짝 진한) 알약 모양으로 그린다.
        // 테두리 선이 줄마다 반복되면 스크롤할 때 끊김이 더 눈에 띄기 때문이다. 포커스일 때만 강조색 테두리를 둔다.
        var surface = UiDraw.SurfaceBehind(this);
        var bounds = new RectangleF(0.5F, 0.5F, Width - 1.5F, Height - 1.5F);
        using (var path = UiDraw.RoundedRect(bounds, bounds.Height / 2))
        {
            var amount = (AppTheme.IsDark ? 0.07F : 0.045F) + (_hovered && Enabled ? 0.04F : 0F);
            using var fill = new SolidBrush(UiDraw.Blend(surface, theme.TextPrimary, amount));
            g.FillPath(fill, path);
            if (Focused)
            {
                using var border = new Pen(theme.Accent, 1.5F);
                g.DrawPath(border, path);
            }
        }

        var text = _selectedIndex >= 0 ? _items[_selectedIndex] : string.Empty;
        var textColor = Enabled ? theme.TextPrimary : theme.TextSecondary;
        // 오른쪽 ⌄와 겹치지 않도록 양쪽 여백을 같게 두고 가운데 정렬한다(고정 글자 칸과 가운데를 맞춘다).
        var side = UiDraw.S(36);
        var textBounds = new Rectangle(side, 0, Width - side * 2, Height);
        TextRenderer.DrawText(g, text, Font, textBounds, textColor,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding);

        using var iconFont = UiDraw.IconFont(7.5F);
        var chevronBounds = new Rectangle(Width - UiDraw.S(36), 0, UiDraw.S(24), Height);
        TextRenderer.DrawText(g, Glyphs.ChevronDown, iconFont, chevronBounds, theme.TextSecondary,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
    }
}
