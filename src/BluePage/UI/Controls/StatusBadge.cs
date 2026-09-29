namespace Microsoft365OfficeWebLauncher.UI.Controls;

internal enum BadgeKind
{
    Neutral,
    Success,
    Warning,
    Danger,
    Accent,
    /// <summary>상태가 아니라 꾸밈용 코랄(아이콘의 노을색). 아이콘 타일에 쓴다.</summary>
    Warm
}

internal static class BadgeColors
{
    public static (Color Back, Color Fore) For(BadgeKind kind)
    {
        var theme = AppTheme.Current;
        return kind switch
        {
            BadgeKind.Success => (theme.SuccessSoft, theme.SuccessSoftText),
            BadgeKind.Warning => (theme.WarningSoft, theme.WarningSoftText),
            BadgeKind.Danger or BadgeKind.Warm => (theme.DangerSoft, theme.DangerSoftText),
            BadgeKind.Accent => (theme.AccentSoft, theme.AccentSoftText),
            _ => (theme.NeutralSoft, theme.TextSecondary)
        };
    }
}

/// <summary>"연결됨", "로그인 필요" 같은 상태를 알약 모양으로 보여 준다. 글자 길이에 맞춰 크기가 바뀐다.</summary>
internal sealed class StatusBadge : Control
{
    private BadgeKind _kind;

    public StatusBadge()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
        Font = new Font("Segoe UI Variable Text Semibold", 8.5F, FontStyle.Bold);
        Margin = new Padding(0);
    }

    public BadgeKind Kind
    {
        get => _kind;
        set { _kind = value; Invalidate(); }
    }

    public void Set(BadgeKind kind, string text)
    {
        _kind = kind;
        Text = text;
        Visible = !string.IsNullOrEmpty(text);
        Invalidate();
    }

    protected override void OnTextChanged(EventArgs e)
    {
        base.OnTextChanged(e);
        var textSize = TextRenderer.MeasureText(Text, Font, Size.Empty, TextFormatFlags.NoPadding);
        Size = new Size(textSize.Width + UiDraw.S(20), textSize.Height + UiDraw.S(8));
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.Clear(UiDraw.SurfaceBehind(this));
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        var (back, fore) = BadgeColors.For(_kind);
        using var path = UiDraw.RoundedRect(new RectangleF(0, 0, Width - 1, Height - 1), Height / 2F);
        using var fill = new SolidBrush(back);
        e.Graphics.FillPath(fill, path);
        TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, fore,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
    }
}

/// <summary>카드 왼쪽의 둥근 아이콘 타일. 아이콘 글꼴 글리프나 글자 하나(예: "M")를 가운데에 그린다.</summary>
internal sealed class IconTile : Control
{
    private readonly bool _useIconFont;
    private BadgeKind _kind;

    public IconTile(string glyphOrLetter, BadgeKind kind, bool useIconFont)
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        _useIconFont = useIconFont;
        _kind = kind;
        Text = glyphOrLetter;
        Size = new Size(UiDraw.S(40), UiDraw.S(40));
        Margin = new Padding(0);
        Font = useIconFont ? UiDraw.IconFont(13F) : new Font("Segoe UI Variable Display Semibold", 13F, FontStyle.Bold);
    }

    public BadgeKind Kind
    {
        get => _kind;
        set { _kind = value; Invalidate(); }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.Clear(UiDraw.SurfaceBehind(this));
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        var (back, fore) = BadgeColors.For(_kind);
        using var path = UiDraw.RoundedRect(new RectangleF(0, 0, Width - 1, Height - 1), UiDraw.S(12));
        using var fill = new SolidBrush(back);
        e.Graphics.FillPath(fill, path);
        var flags = TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix;
        if (!_useIconFont)
        {
            flags |= TextFormatFlags.SingleLine;
        }
        TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, fore, flags);
    }
}
