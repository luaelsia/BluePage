using System.Drawing.Drawing2D;
using System.Globalization;

namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>
/// 사이드바 아래쪽 정보 영역. 앱 아이콘에서 따 온 고래(몸통, 밝은 등선, 말린 꼬리, 그라데이션 꼬리 조각)를
/// 사이드바 색에 청록을 조금 섞은 색으로 옅게 그리고, 고래 아래는 몸통 색으로 바닥까지 채운다.
/// 버전과 연락처 글자는 고래 몸통 안에 들어간다. 안쪽 글자 컨트롤은 배경을 투명하게 둔다.
/// </summary>
internal sealed class WhaleFooter : Panel
{
    /// <summary>고래 바닥에서 컨트롤 바닥까지 몸통 색으로 채우는 높이(96 DPI 기준).</summary>
    public int BodyBelowWhale { get; set; } = UiDraw.S(62);

    public WhaleFooter()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaintBackground(PaintEventArgs e)
    {
        var theme = AppTheme.Current;
        var dark = AppTheme.IsDark;
        var sidebar = theme.Sidebar;
        var g = e.Graphics;
        g.Clear(sidebar);
        g.SmoothingMode = SmoothingMode.AntiAlias;

        var body = UiDraw.Blend(sidebar, theme.Accent, dark ? 0.14F : 0.10F);
        var ribbon = UiDraw.Blend(sidebar, theme.Accent, dark ? 0.30F : 0.22F);
        var finBottom = UiDraw.Blend(sidebar, theme.Accent, dark ? 0.20F : 0.14F);

        var scale = Width / IconShapes.BaseWidth;
        var whaleBottom = Height - BodyBelowWhale;

        using (var fill = new SolidBrush(body))
        {
            g.FillRectangle(fill, 0, whaleBottom, Width, Height - whaleBottom);
        }

        var state = g.Save();
        g.TranslateTransform(0, whaleBottom + 0.5F);
        g.ScaleTransform(scale, scale);
        using (var bodyPath = SvgPath.Parse(IconShapes.WhaleBody))
        using (var ribbonPath = SvgPath.Parse(IconShapes.WhaleRibbon))
        using (var finPath = SvgPath.Parse(IconShapes.WhaleFin))
        using (var bodyBrush = new SolidBrush(body))
        using (var ribbonBrush = new SolidBrush(ribbon))
        using (var finBrush = new LinearGradientBrush(
                   new PointF(0, IconShapes.WhaleFinTop - 0.5F), new PointF(0, 0.5F), ribbon, finBottom))
        {
            g.FillPath(bodyBrush, bodyPath);
            g.FillPath(ribbonBrush, ribbonPath);
            g.FillPath(finBrush, finPath);
        }
        g.Restore(state);
    }
}

/// <summary>
/// 사이드바 오른쪽 위의 접힌 모서리. 사이드바 색과 오른쪽 내용 영역 색을 직접 칠하므로 투명 처리가 필요 없다.
/// 접힌 선이 라이트 테마에서도 보이도록 가장자리 선과 옅은 그림자를 함께 그린다.
/// </summary>
internal sealed class FoldCorner : Control
{
    public FoldCorner(float scale)
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        _scale = scale;
        Size = new Size((int)Math.Ceiling(IconShapes.FoldWidth * scale), (int)Math.Ceiling(IconShapes.FoldHeight * scale));
        Margin = new Padding(0);
        TabStop = false;
    }

    private readonly float _scale;

    protected override void OnPaint(PaintEventArgs e)
    {
        var theme = AppTheme.Current;
        var dark = AppTheme.IsDark;
        var g = e.Graphics;
        g.Clear(theme.Sidebar);
        g.SmoothingMode = SmoothingMode.AntiAlias;

        var top = dark ? UiDraw.Blend(theme.Sidebar, Color.FromArgb(243, 230, 195), 0.22F) : Color.FromArgb(255, 247, 230);
        var under = UiDraw.Blend(theme.Sidebar, Color.FromArgb(233, 184, 100), dark ? 0.25F : 0.45F);
        var shadow = dark ? UiDraw.Blend(theme.Sidebar, Color.Black, 0.30F) : Color.FromArgb(233, 220, 198);
        var crease = UiDraw.Blend(theme.Sidebar, Color.FromArgb(220, 200, 162), dark ? 0.34F : 0.55F);

        g.TranslateTransform(Width, 0);
        g.ScaleTransform(_scale, _scale);
        using var cut = SvgPath.Parse(IconShapes.FoldCut);
        using var underPath = SvgPath.Parse(IconShapes.FoldUnder);
        using var topPath = SvgPath.Parse(IconShapes.FoldTop);
        using var creasePath = SvgPath.Parse(IconShapes.FoldCrease);

        using (var brush = new SolidBrush(theme.Background))
        {
            g.FillPath(brush, cut);
        }

        var state = g.Save();
        g.TranslateTransform(1.4F, 1.8F);
        using (var brush = new SolidBrush(Color.FromArgb(dark ? 150 : 200, shadow)))
        {
            g.FillPath(brush, topPath);
        }
        g.Restore(state);

        using (var brush = new SolidBrush(under))
        {
            g.FillPath(brush, underPath);
        }
        using (var brush = new SolidBrush(top))
        {
            g.FillPath(brush, topPath);
        }
        using (var pen = new Pen(crease, (dark ? 0.8F : 1.1F) / _scale * UiDraw.S(100) / 100F))
        {
            g.DrawPath(pen, creasePath);
        }
    }
}

/// <summary>"M x,y L x,y ... Z" 형태의 간단한 SVG 경로만 읽는다(IconShapes 전용).</summary>
internal static class SvgPath
{
    private static readonly Dictionary<string, PointF[][]> Cache = new();

    public static GraphicsPath Parse(string data)
    {
        if (!Cache.TryGetValue(data, out var figures))
        {
            figures = ParseFigures(data);
            Cache[data] = figures;
        }

        var path = new GraphicsPath();
        foreach (var figure in figures)
        {
            if (figure.Length < 2)
            {
                continue;
            }
            path.StartFigure();
            path.AddLines(figure);
        }
        // 닫힌 도형(Z)이 있는 경로만 닫는다. 접힌 선처럼 열린 선은 그대로 둔다.
        if (data.TrimEnd().EndsWith('Z'))
        {
            path.CloseAllFigures();
        }
        return path;
    }

    private static PointF[][] ParseFigures(string data)
    {
        var figures = new List<PointF[]>();
        var current = new List<PointF>();
        foreach (var token in data.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (token == "Z")
            {
                continue;
            }
            var text = token;
            if (text[0] == 'M')
            {
                if (current.Count > 0) figures.Add(current.ToArray());
                current = new List<PointF>();
                text = text[1..];
            }
            else if (text[0] == 'L')
            {
                text = text[1..];
            }
            var parts = text.Split(',');
            current.Add(new PointF(
                float.Parse(parts[0], CultureInfo.InvariantCulture),
                float.Parse(parts[1], CultureInfo.InvariantCulture)));
        }
        if (current.Count > 0) figures.Add(current.ToArray());
        return figures.ToArray();
    }
}
