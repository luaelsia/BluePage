using Microsoft365OfficeWebLauncher.UI.Controls;
using System.Drawing.Drawing2D;

namespace Microsoft365OfficeWebLauncher.UI;

internal sealed class ModernButton : Button
{
    private bool _hovered;
    private bool _pressed;
    private bool _isPrimary;

    /// <summary>화면마다 하나만 두는 주요 동작 버튼. 강조색으로 채운다.</summary>
    public bool IsPrimary
    {
        get => _isPrimary;
        set { _isPrimary = value; Invalidate(); }
    }

    public ModernButton()
    {
        SetStyle(
            ControlStyles.UserPaint |
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.ResizeRedraw,
            true);

        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        UseVisualStyleBackColor = false;
        Cursor = Cursors.Hand;
        Height = UiDraw.S(34);
        Padding = UiDraw.S(14, 0, 14, 0);
        MinimumSize = new Size(UiDraw.S(76), UiDraw.S(34));
    }

    protected override void OnMouseEnter(EventArgs e)
    {
        _hovered = true;
        Invalidate();
        base.OnMouseEnter(e);
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        _hovered = false;
        _pressed = false;
        Invalidate();
        base.OnMouseLeave(e);
    }

    protected override void OnMouseDown(MouseEventArgs mevent)
    {
        if (mevent.Button == MouseButtons.Left)
        {
            _pressed = true;
            Invalidate();
        }
        base.OnMouseDown(mevent);
    }

    protected override void OnMouseUp(MouseEventArgs mevent)
    {
        _pressed = false;
        Invalidate();
        base.OnMouseUp(mevent);
    }

    protected override void OnEnabledChanged(EventArgs e)
    {
        Invalidate();
        base.OnEnabledChanged(e);
    }

    protected override void OnGotFocus(EventArgs e)
    {
        Invalidate();
        base.OnGotFocus(e);
    }

    protected override void OnLostFocus(EventArgs e)
    {
        Invalidate();
        base.OnLostFocus(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var theme = AppTheme.Current;
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        var surface = Parent?.BackColor ?? theme.Background;
        e.Graphics.Clear(surface);

        var bounds = Rectangle.Inflate(ClientRectangle, -1, -1);
        // 알약 모양: 높이의 절반을 반지름으로 쓴다.
        using var path = CreateRoundedPath(bounds, Math.Max(1, bounds.Height / 2));

        Color background;
        Color textColor;
        if (_isPrimary)
        {
            background = _pressed ? Blend(theme.Accent, theme.TextPrimary, 0.15F)
                : _hovered ? Blend(theme.Accent, theme.CardBackground, 0.12F)
                : theme.Accent;
            textColor = theme.OnAccent;
        }
        else
        {
            background = _pressed ? Blend(theme.ButtonHover, theme.TextPrimary, 0.08F)
                : _hovered ? theme.ButtonHover
                : theme.ButtonBackground;
            textColor = theme.TextPrimary;
        }
        if (!Enabled)
        {
            background = Blend(background, surface, 0.5F);
            textColor = _isPrimary ? Blend(theme.OnAccent, background, 0.3F) : theme.TextSecondary;
        }

        using var fill = new SolidBrush(background);
        e.Graphics.FillPath(fill, path);

        if (!_isPrimary || Focused)
        {
            var borderColor = Focused ? theme.Accent : theme.ButtonBorder;
            using var border = new Pen(borderColor, Focused ? 1.5F : 1F);
            e.Graphics.DrawPath(border, path);
        }

        TextRenderer.DrawText(
            e.Graphics,
            Text,
            Font,
            bounds,
            textColor,
            TextFormatFlags.HorizontalCenter |
            TextFormatFlags.VerticalCenter |
            TextFormatFlags.EndEllipsis |
            TextFormatFlags.NoPrefix);
    }

    private static GraphicsPath CreateRoundedPath(Rectangle bounds, int radius)
    {
        var path = new GraphicsPath();
        var diameter = radius * 2;
        path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
        path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
        path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        return path;
    }

    private static Color Blend(Color first, Color second, float amount) => Color.FromArgb(
        (int)(first.R + (second.R - first.R) * amount),
        (int)(first.G + (second.G - first.G) * amount),
        (int)(first.B + (second.B - first.B) * amount));
}
