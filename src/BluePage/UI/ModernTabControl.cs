namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>WinForms 기본 탭의 운영체제별 흰 배경/테두리를 사용하지 않는 Blue Page 전용 탭 컨테이너.</summary>
public sealed class ModernTabControl : UserControl
{
    private readonly FlowLayoutPanel _headers;
    private readonly Panel _contentHost;
    private readonly List<(Button Header, Control Content)> _tabs = new();
    private int _selectedIndex = -1;
    private ThemePalette _theme = AppTheme.Current;

    public ModernTabControl()
    {
        DoubleBuffered = true;
        Padding = new Padding(0);

        _headers = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 38,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0),
            Margin = new Padding(0)
        };

        _contentHost = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(20),
            Margin = new Padding(0)
        };
        _contentHost.Paint += PaintContentBorder;

        Controls.Add(_contentHost);
        Controls.Add(_headers);
    }

    public void AddTab(string title, Control content)
    {
        var index = _tabs.Count;
        var header = new Button
        {
            Text = title,
            Width = 112,
            Height = 38,
            FlatStyle = FlatStyle.Flat,
            TabStop = true,
            Margin = new Padding(0),
            Padding = new Padding(0),
            UseVisualStyleBackColor = false,
            AccessibleName = $"{title} 탭"
        };
        header.Click += (_, _) => SelectTab(index);
        header.Paint += (_, e) => PaintHeader(e.Graphics, header, index == _selectedIndex);

        content.Dock = DockStyle.Fill;
        content.Visible = false;
        _headers.Controls.Add(header);
        _contentHost.Controls.Add(content);
        _tabs.Add((header, content));

        if (_selectedIndex < 0)
        {
            SelectTab(0);
        }
        ApplyTheme(AppTheme.Current);
    }

    public void ApplyTheme(ThemePalette theme)
    {
        _theme = theme;
        BackColor = theme.Background;
        ForeColor = theme.TextPrimary;
        _headers.BackColor = theme.Background;
        _contentHost.BackColor = theme.Background;
        _contentHost.ForeColor = theme.TextPrimary;
        foreach (var (header, _) in _tabs)
        {
            header.BackColor = theme.Background;
            header.ForeColor = theme.TextPrimary;
            header.FlatAppearance.BorderSize = 0;
            header.Invalidate();
        }
        _contentHost.Invalidate();
        Invalidate();
    }

    private void SelectTab(int index)
    {
        if (index < 0 || index >= _tabs.Count)
        {
            return;
        }
        for (var i = 0; i < _tabs.Count; i++)
        {
            _tabs[i].Content.Visible = i == index;
        }
        _selectedIndex = index;
        _tabs[index].Content.BringToFront();
        foreach (var (header, _) in _tabs) header.Invalidate();
    }

    private void PaintHeader(Graphics graphics, Button header, bool selected)
    {
        graphics.Clear(selected ? _theme.CardBackground : _theme.Background);
        var borderColor = selected ? _theme.Accent : _theme.ButtonBorder;
        using var borderPen = new Pen(borderColor);
        graphics.DrawRectangle(borderPen, 0, 0, header.Width - 1, header.Height - 1);
        if (selected)
        {
            using var accentPen = new Pen(_theme.Accent, 3);
            graphics.DrawLine(accentPen, 1, header.Height - 2, header.Width - 2, header.Height - 2);
        }
        TextRenderer.DrawText(
            graphics,
            header.Text,
            header.Font,
            header.ClientRectangle,
            selected ? _theme.TextPrimary : _theme.TextSecondary,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
    }

    private void PaintContentBorder(object? sender, PaintEventArgs e)
    {
        using var pen = new Pen(_theme.ButtonBorder);
        e.Graphics.DrawRectangle(pen, 0, 0, _contentHost.Width - 1, _contentHost.Height - 1);
    }
}
