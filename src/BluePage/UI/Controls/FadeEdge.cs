using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>
/// 스크롤 화면 위나 아래 가장자리의 흐림 띠. 가장자리 쪽은 배경색이 불투명하고 안쪽으로 갈수록 투명해진다.
/// 레이어드 자식 창(Windows 8 이상, app.manifest의 Windows 10 선언 필요)이라 아래 카드가 실제로 비쳐 보인다.
/// 클릭은 통과시킨다(띠 밑의 버튼도 그대로 눌린다).
/// </summary>
internal sealed class FadeEdge : Control
{
    private const int WS_EX_LAYERED = 0x00080000;
    private const int WS_EX_TRANSPARENT = 0x00000020;
    private const int WM_NCHITTEST = 0x0084;
    private const int HTTRANSPARENT = -1;

    private readonly bool _top;

    /// <summary>이 띠가 덮고 있는 스크롤 화면. 띠 위에서 돌린 휠을 이 화면으로 보낸다.</summary>
    public ScrollableControl? Owner { get; init; }

    public FadeEdge(bool top)
    {
        _top = top;
        TabStop = false;
        Visible = false;
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var parameters = base.CreateParams;
            parameters.ExStyle |= WS_EX_LAYERED | WS_EX_TRANSPARENT;
            return parameters;
        }
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_NCHITTEST)
        {
            m.Result = (IntPtr)HTTRANSPARENT;
            return;
        }
        base.WndProc(ref m);
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Render();
    }

    protected override void OnSizeChanged(EventArgs e)
    {
        base.OnSizeChanged(e);
        Render();
    }

    /// <summary>테마가 바뀌면 ThemeApplier가 Invalidate를 부른다. 레이어드 창은 직접 다시 만들어 올린다.</summary>
    protected override void OnInvalidated(InvalidateEventArgs e)
    {
        base.OnInvalidated(e);
        Render();
    }

    /// <summary>배경색 그라데이션을 픽셀 단위 투명도로 창에 올린다.</summary>
    private void Render()
    {
        if (!IsHandleCreated || Width <= 0 || Height <= 0)
        {
            return;
        }

        var color = AppTheme.Current.Background;
        using var bitmap = new Bitmap(Width, Height, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bitmap))
        using (var brush = new LinearGradientBrush(
                   new Rectangle(0, 0, Width, Height),
                   Color.FromArgb(255, color),
                   Color.FromArgb(0, color),
                   _top ? 90F : 270F))
        {
            g.Clear(Color.Transparent);
            g.FillRectangle(brush, 0, 0, Width, Height);
        }

        var screenDc = GetDC(IntPtr.Zero);
        var memoryDc = CreateCompatibleDC(screenDc);
        var hBitmap = bitmap.GetHbitmap(Color.FromArgb(0));
        var previous = SelectObject(memoryDc, hBitmap);
        try
        {
            var size = new NativeSize(Width, Height);
            var source = new NativePoint();
            var blend = new BlendFunction { BlendOp = 0, BlendFlags = 0, SourceConstantAlpha = 255, AlphaFormat = 1 };
            UpdateLayeredWindow(Handle, screenDc, IntPtr.Zero, ref size, memoryDc, ref source, 0, ref blend, 2);
        }
        finally
        {
            SelectObject(memoryDc, previous);
            DeleteObject(hBitmap);
            DeleteDC(memoryDc);
            ReleaseDC(IntPtr.Zero, screenDc);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeSize
    {
        public int Width;
        public int Height;
        public NativeSize(int width, int height) { Width = width; Height = height; }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    private struct BlendFunction
    {
        public byte BlendOp;
        public byte BlendFlags;
        public byte SourceConstantAlpha;
        public byte AlphaFormat;
    }

    [DllImport("user32.dll")]
    private static extern bool UpdateLayeredWindow(IntPtr hwnd, IntPtr hdcDst, IntPtr pptDst, ref NativeSize psize,
        IntPtr hdcSrc, ref NativePoint pptSrc, int crKey, ref BlendFunction pblend, int dwFlags);

    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    private static extern IntPtr SelectObject(IntPtr hdc, IntPtr obj);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr obj);
}

/// <summary>
/// 한 스크롤 화면의 위아래 흐림 띠 한 쌍. 위로 더 스크롤할 내용이 있을 때만 위 띠를,
/// 아래로 더 스크롤할 내용이 있을 때만 아래 띠를 보인다. 화면이 숨겨지면 둘 다 숨긴다.
/// 띠를 만들 수 없는 환경이면(레이어드 자식 창 미지원) 조용히 띠 없이 둔다.
/// </summary>
internal sealed class ScrollFades
{
    private readonly ScrollableControl _panel;
    private readonly FadeEdge? _top;
    private readonly FadeEdge? _bottom;

    public ScrollFades(ScrollableControl panel, Control host, SmoothScroller scroller, int height)
    {
        _panel = panel;
        try
        {
            _top = new FadeEdge(top: true) { Height = height, Owner = panel };
            _bottom = new FadeEdge(top: false) { Height = height, Owner = panel };
            host.Controls.Add(_top);
            host.Controls.Add(_bottom);
            _top.BringToFront();
            _bottom.BringToFront();
            _ = _top.Handle;
            _ = _bottom.Handle;
        }
        catch (System.ComponentModel.Win32Exception)
        {
            // 레이어드 자식 창을 만들 수 없으면 흐림 없이 둔다.
            if (_top is not null) host.Controls.Remove(_top);
            if (_bottom is not null) host.Controls.Remove(_bottom);
            _top = null;
            _bottom = null;
            return;
        }

        panel.Scroll += (_, _) => Update(-panel.AutoScrollPosition.Y);
        panel.Resize += (_, _) => Update(-panel.AutoScrollPosition.Y);
        panel.Layout += (_, _) => Update(-panel.AutoScrollPosition.Y);
        panel.VisibleChanged += (_, _) => Update(-panel.AutoScrollPosition.Y);
        scroller.Scrolled += (scrolled, position) =>
        {
            if (scrolled == panel)
            {
                Update(position);
            }
        };
    }

    private void Update(float position)
    {
        if (_top is null || _bottom is null)
        {
            return;
        }

        var bounds = new Rectangle(_panel.Left, _panel.Top, _panel.ClientSize.Width, _panel.ClientSize.Height);
        _top.SetBounds(bounds.Left, bounds.Top, bounds.Width, _top.Height);
        _bottom.SetBounds(bounds.Left, bounds.Bottom - _bottom.Height, bounds.Width, _bottom.Height);

        var max = SmoothScroller.MaxScroll(_panel);
        var shown = _panel.Visible && max > 0;
        _top.Visible = shown && position > 0.5F;
        _bottom.Visible = shown && position < max - 0.5F;
    }
}
