namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>
/// 체크박스 대신 쓰는 켜기/끄기 스위치. CheckBox를 상속하므로 Checked, CheckedChanged를 그대로 쓴다.
/// 글자는 그리지 않는다. 설명은 행의 왼쪽 라벨이 맡는다.
/// </summary>
internal sealed class ToggleSwitch : CheckBox
{
    private bool _hovered;

    public ToggleSwitch()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        AutoSize = false;
        Size = new Size(UiDraw.S(42), UiDraw.S(24));
        Cursor = Cursors.Hand;
        Margin = new Padding(0);
        Text = string.Empty;
    }

    protected override void OnMouseEnter(EventArgs eventargs) { _hovered = true; Invalidate(); base.OnMouseEnter(eventargs); }
    protected override void OnMouseLeave(EventArgs eventargs) { _hovered = false; Invalidate(); base.OnMouseLeave(eventargs); }
    protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }

    protected override void OnPaint(PaintEventArgs pevent)
    {
        var theme = AppTheme.Current;
        var g = pevent.Graphics;
        g.Clear(UiDraw.SurfaceBehind(this));
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        var track = new RectangleF(1, 1, Width - 3, Height - 3);
        var trackColor = Checked ? theme.Accent : theme.ButtonBorder;
        if (_hovered && Enabled)
        {
            trackColor = UiDraw.Blend(trackColor, theme.TextPrimary, 0.08F);
        }
        if (!Enabled)
        {
            trackColor = UiDraw.Blend(trackColor, UiDraw.SurfaceBehind(this), 0.55F);
        }

        using (var path = UiDraw.RoundedRect(track, track.Height / 2))
        using (var fill = new SolidBrush(trackColor))
        {
            g.FillPath(fill, path);
        }

        var knobSize = track.Height - UiDraw.S(6);
        var knobX = Checked ? track.Right - knobSize - UiDraw.S(3) : track.Left + UiDraw.S(3);
        var knobColor = Checked ? theme.OnAccent : theme.CardBackground;
        if (!Enabled)
        {
            knobColor = UiDraw.Blend(knobColor, UiDraw.SurfaceBehind(this), 0.35F);
        }
        using var knob = new SolidBrush(knobColor);
        g.FillEllipse(knob, knobX, track.Top + UiDraw.S(3), knobSize, knobSize);

        if (Focused && ShowFocusCues)
        {
            using var focus = new Pen(theme.Accent, 1.5F);
            using var path = UiDraw.RoundedRect(new RectangleF(0.5F, 0.5F, Width - 2, Height - 2), (Height - 2) / 2F);
            g.DrawPath(focus, path);
        }
    }
}
