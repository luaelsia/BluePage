using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>
/// 마우스 휠 스크롤을 부드럽게 만든다.
/// WinForms는 라벨, 버튼, 카드가 모두 별도 창이라 스크롤 위치를 바꿀 때마다 수십 개의 창을 옮기고 다시 그린다.
/// 그래서 움직이는 동안은 화면 내용을 이미지 한 장으로 찍어 보이는 부분만 복사해 그리고,
/// 멈춘 뒤 실제 스크롤 위치를 한 번만 바꾼다.
/// 움직임은 용수철(임계 감쇠)로 목표를 따라간다. 휠이 연달아 들어와도 속도는 그대로 이어지고 목표만 바뀌어서
/// 칸마다 속도가 튀지 않는다.
/// 프레임은 타이머가 아니라 화면 합성 주기(DwmFlush)에 맞춰 그려서 모니터 갱신과 박자가 어긋나지 않는다.
/// 마우스가 스크롤 화면 위에 잠시 멈춰 있으면 이미지를 미리 찍어 두어, 휠을 돌리는 순간 바로 움직인다.
/// DataGridView(동기화 검토 표)는 줄 단위로만 움직일 수 있어서 한 칸에 한 줄씩 넘긴다.
/// 드롭다운 메뉴처럼 다른 창 위에서 돌린 휠은 건드리지 않는다.
/// </summary>
internal sealed class SmoothScroller : IMessageFilter, IDisposable
{
    private const int WM_MOUSEWHEEL = 0x020A;
    private const int WM_MOUSEMOVE = 0x0200;
    private const int WM_LBUTTONDOWN = 0x0201;
    private const int WM_RBUTTONDOWN = 0x0204;
    private const int WM_KEYDOWN = 0x0100;

    /// <summary>용수철 세기(초당 라디안). 한 칸(90px)이 약 0.2초 안에 멈춘다.</summary>
    private const double SpringOmega = 30;
    /// <summary>움직임이 멈춘 뒤 이미지를 더 유지하는 시간. 이 안에 다음 휠이 오면 찍기와 바꾸기 없이 바로 이어간다.</summary>
    private const double LingerMs = 500;
    /// <summary>마우스가 이만큼 가만히 있으면 이미지를 미리 찍는다.</summary>
    private const int PrecaptureDelayMs = 150;
    /// <summary>미리 찍은 이미지를 믿고 쓰는 시간. 지나면 휠을 돌릴 때 새로 찍는다.</summary>
    private const double PrecaptureMaxAgeMs = 2000;

    private sealed class ScrollState
    {
        public double Position;
        public double Velocity;
        public double Target;
        public long LastTicks;
        public bool Animating;
        /// <summary>움직임은 멈췄고 이미지를 유지하는 중(Animating도 true).</summary>
        public bool Lingering;
        public long LingerStartTicks;
        public SnapshotOverlay? Overlay;
        /// <summary>미리 찍어 둔 이미지(마우스가 멈춰 있을 때 찍음).</summary>
        public CapturedContent? Precaptured;
    }

    private readonly Dictionary<ScrollableControl, ScrollState> _panels = new();
    /// <summary>화면 합성 주기를 기다릴 수 없을 때 쓰는 예비 타이머.</summary>
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 15 };
    private readonly System.Windows.Forms.Timer _precaptureTimer = new() { Interval = PrecaptureDelayMs };
    private ScrollableControl? _hoverPanel;
    private SynchronizationContext? _uiContext;
    private volatile bool _pumpRunning;
    private int _stepQueued;

    /// <summary>움직이는 동안 보이는 스크롤 위치가 바뀔 때(끝날 때 포함). 흐림 띠 표시에 쓴다.</summary>
    public event Action<ScrollableControl, float>? Scrolled;

    public SmoothScroller()
    {
        _timer.Tick += (_, _) => Step();
        _precaptureTimer.Tick += (_, _) =>
        {
            _precaptureTimer.Stop();
            Precapture();
        };
    }

    /// <summary>휠 한 칸에 움직이는 거리(96 DPI 기준 픽셀).</summary>
    public int PixelsPerNotch { get; set; } = UiDraw.S(90);

    public void Register(ScrollableControl panel)
    {
        var state = new ScrollState();
        _panels[panel] = state;
        panel.Disposed += (_, _) => _panels.Remove(panel);
        // 이미지를 유지하는 중에 화면이 바뀌거나 크기가 바뀌면 바로 실제 화면으로 돌아간다.
        panel.VisibleChanged += (_, _) => { Finish(panel, state); DropPrecaptured(state); };
        panel.Resize += (_, _) => { Finish(panel, state); DropPrecaptured(state); };
        // 배치가 바뀌면(내용이 늘거나 줄면) 미리 찍은 이미지는 더 이상 맞지 않는다.
        panel.Layout += (_, _) => { if (!state.Animating) DropPrecaptured(state); };
    }

    /// <summary>스크롤할 수 있는 최대 위치. 스크롤바가 실제로 쓰는 범위라 스크롤바를 끌었을 때의 끝과 같다.</summary>
    public static int MaxScroll(ScrollableControl panel)
    {
        var bar = panel.VerticalScroll;
        return bar.Visible ? Math.Max(0, bar.Maximum - bar.LargeChange + 1) : 0;
    }

    public bool PreFilterMessage(ref Message m)
    {
        switch (m.Msg)
        {
            case WM_MOUSEMOVE:
                OnMouseMoved(m.HWnd);
                return false;
            case WM_LBUTTONDOWN or WM_RBUTTONDOWN or WM_KEYDOWN:
                // 누르거나 입력하면 화면 내용이 바뀔 수 있으니 미리 찍은 이미지를 버리고, 잠시 뒤 다시 찍는다.
                foreach (var panelState in _panels.Values)
                {
                    DropPrecaptured(panelState);
                }
                _precaptureTimer.Stop();
                _precaptureTimer.Start();
                return false;
            case WM_MOUSEWHEEL:
                break;
            default:
                return false;
        }

        var target = Control.FromChildHandle(WindowFromPoint(Cursor.Position));
        if (target is null)
        {
            return false;
        }

        var delta = (short)((long)m.WParam >> 16);
        for (var control = target; control is not null; control = control.Parent)
        {
            // 움직이는 중에는 이미지가 패널을 덮고 있으므로 이미지의 주인 패널로 돌린다.
            if (control is SnapshotOverlay overlay && _panels.TryGetValue(overlay.Owner, out var overlayState))
            {
                return ScrollPanel(overlay.Owner, overlayState, delta);
            }
            if (control is FadeEdge { Owner: { } fadeOwner } && _panels.TryGetValue(fadeOwner, out var fadeState))
            {
                return ScrollPanel(fadeOwner, fadeState, delta);
            }
            if (control is DataGridView grid)
            {
                return ScrollGridByRow(grid, delta);
            }
            if (control is ScrollableControl panel && _panels.TryGetValue(panel, out var state))
            {
                return ScrollPanel(panel, state, delta);
            }
        }
        return false;
    }

    private bool ScrollPanel(ScrollableControl panel, ScrollState state, int delta)
    {
        var max = MaxScroll(panel);
        if (max == 0)
        {
            return false;
        }

        if (!state.Animating)
        {
            // 멈춰 있으면(스크롤바를 끌었을 수 있으니) 실제 위치에서 정지 상태로 출발한다.
            var scroll = -panel.AutoScrollPosition.Y;
            state.Position = scroll;
            state.Velocity = 0;
            state.Target = scroll;

            var captured = TakePrecaptured(panel, state, scroll) ?? CapturedContent.Capture(panel, scroll);
            state.Overlay = captured is null ? null : SnapshotOverlay.TryCreate(panel, captured, scroll, () => Finish(panel, state));
            state.LastTicks = Stopwatch.GetTimestamp();
        }
        else if (state.Lingering)
        {
            // 유지 중이던 이미지에서 이어서 움직인다(찍기 없음). 멈춰 있던 시간은 계산에 넣지 않는다.
            state.LastTicks = Stopwatch.GetTimestamp();
        }

        // 속도는 건드리지 않고 목표만 옮긴다. 그래서 연달아 굴려도 칸마다 속도가 튀지 않는다.
        state.Target = Math.Clamp(state.Target - delta / 120.0 * PixelsPerNotch, 0, max);
        state.Animating = true;
        state.Lingering = false;
        StartTimer();
        return true;
    }

    /// <summary>임계 감쇠 용수철의 정확한 해로 dt만큼 진행한다(프레임 간격이 흔들려도 결과가 안정적이다).</summary>
    private static void Advance(ScrollState state, double dt)
    {
        var offset = state.Position - state.Target;
        var decay = Math.Exp(-SpringOmega * dt);
        var temp = (state.Velocity + SpringOmega * offset) * dt;
        state.Position = state.Target + (offset + temp) * decay;
        state.Velocity = (state.Velocity - SpringOmega * temp) * decay;
    }

    private void Step()
    {
        var now = Stopwatch.GetTimestamp();
        var moving = false;
        foreach (var (panel, state) in _panels.ToList())
        {
            if (!state.Animating || panel.IsDisposed)
            {
                continue;
            }

            if (state.Lingering)
            {
                // 움직임은 멈췄다. 유지 시간이 지나면 실제 화면으로 바꾼다(그 전까지는 아무것도 하지 않는다).
                if (Ms(state.LingerStartTicks, now) >= LingerMs)
                {
                    Finish(panel, state);
                }
                else
                {
                    moving = true;
                }
                continue;
            }

            // 한 프레임이 너무 길게 밀렸을 때 한꺼번에 튀지 않도록 dt 상한을 둔다.
            var dt = Math.Min(Ms(state.LastTicks, now) / 1000.0, 0.05);
            state.LastTicks = now;
            Advance(state, dt);

            var settled = Math.Abs(state.Position - state.Target) < 0.5 && Math.Abs(state.Velocity) < 6;
            if (settled)
            {
                state.Position = state.Target;
                state.Velocity = 0;
            }

            if (state.Overlay is { } overlay)
            {
                overlay.ScrollTo((float)state.Position);
            }
            else
            {
                // 이미지를 만들지 못했으면 실제 위치를 직접 옮긴다.
                panel.AutoScrollPosition = new Point(0, (int)Math.Round(state.Position));
            }
            Scrolled?.Invoke(panel, (float)state.Position);

            if (settled)
            {
                if (state.Overlay is not null)
                {
                    // 목표에 닿았다. 이미지를 잠시 유지한다.
                    state.Lingering = true;
                    state.LingerStartTicks = now;
                    moving = true;
                }
                else
                {
                    Finish(panel, state);
                }
                continue;
            }
            moving = true;
        }

        if (!moving)
        {
            StopTimer();
        }
    }

    /// <summary>실제 스크롤 위치를 목표로 한 번 옮기고, 실제 화면을 다 그린 뒤 이미지를 치운다.</summary>
    private void Finish(ScrollableControl panel, ScrollState state)
    {
        if (!state.Animating)
        {
            return;
        }
        state.Animating = false;
        state.Lingering = false;
        state.Velocity = 0;

        panel.AutoScrollPosition = new Point(0, (int)Math.Round(state.Target));
        if (state.Overlay is { } overlay)
        {
            state.Overlay = null;
            overlay.Parent?.Controls.Remove(overlay);
            overlay.Dispose();
            // 이미지가 치워진 자리를 안쪽 컨트롤까지 한 번에 바로 그려서 빈 화면이 비치지 않게 한다.
            RedrawWindow(panel.Handle, IntPtr.Zero, IntPtr.Zero, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        }
        Scrolled?.Invoke(panel, -panel.AutoScrollPosition.Y);

        if (!_panels.Values.Any(s => s.Animating))
        {
            StopTimer();
        }
    }

    private static bool ScrollGridByRow(DataGridView grid, int delta)
    {
        var first = grid.FirstDisplayedScrollingRowIndex;
        if (first < 0)
        {
            return false;
        }

        var notches = Math.Max(1, Math.Abs(delta) / 120);
        var index = first;
        for (var i = 0; i < notches; i++)
        {
            var next = delta < 0
                ? grid.Rows.GetNextRow(index, DataGridViewElementStates.Visible)
                : grid.Rows.GetPreviousRow(index, DataGridViewElementStates.Visible);
            if (next < 0)
            {
                break;
            }
            index = next;
        }

        if (index != first)
        {
            try
            {
                grid.FirstDisplayedScrollingRowIndex = index;
            }
            catch (InvalidOperationException)
            {
                // 표가 다시 그려지는 중이면 이번 칸은 건너뛴다.
            }
        }
        return true;
    }

    private void OnMouseMoved(IntPtr hwnd)
    {
        ScrollableControl? panel = null;
        for (var control = Control.FromChildHandle(hwnd); control is not null; control = control.Parent)
        {
            if (control is ScrollableControl candidate && _panels.ContainsKey(candidate))
            {
                panel = candidate;
                break;
            }
        }

        _hoverPanel = panel;
        _precaptureTimer.Stop();
        if (panel is not null)
        {
            _precaptureTimer.Start();
        }
    }

    /// <summary>마우스가 멈춘 스크롤 화면의 이미지를 미리 찍어 둔다(이미 있고 아직 쓸 만하면 그대로 둔다).</summary>
    private void Precapture()
    {
        if (_hoverPanel is not { IsDisposed: false, Visible: true } panel ||
            !_panels.TryGetValue(panel, out var state) || state.Animating || MaxScroll(panel) == 0)
        {
            return;
        }

        var scroll = -panel.AutoScrollPosition.Y;
        if (state.Precaptured is { } existing && existing.IsUsable(panel, scroll, PrecaptureMaxAgeMs))
        {
            return;
        }
        DropPrecaptured(state);
        state.Precaptured = CapturedContent.Capture(panel, scroll);
    }

    private static CapturedContent? TakePrecaptured(ScrollableControl panel, ScrollState state, int scroll)
    {
        var captured = state.Precaptured;
        state.Precaptured = null;
        if (captured is not null && captured.IsUsable(panel, scroll, PrecaptureMaxAgeMs))
        {
            return captured;
        }
        captured?.Dispose();
        return null;
    }

    private static void DropPrecaptured(ScrollState state)
    {
        state.Precaptured?.Dispose();
        state.Precaptured = null;
    }

    /// <summary>
    /// 프레임을 그리기 시작한다. 별도 스레드가 화면 합성 주기(DwmFlush)를 기다렸다가 UI 스레드에 한 프레임씩 그리라고 알린다.
    /// UI 스레드 문맥을 얻지 못하면 예비 타이머로 그린다.
    /// </summary>
    private void StartTimer()
    {
        _uiContext ??= SynchronizationContext.Current;
        if (_uiContext is null)
        {
            _timer.Start();
            return;
        }
        if (_pumpRunning)
        {
            return;
        }

        _pumpRunning = true;
        var context = _uiContext;
        var pump = new Thread(() =>
        {
            while (_pumpRunning)
            {
                var waitStart = Stopwatch.GetTimestamp();
                var composed = DwmFlush() == 0;
                // 창이 숨겨졌거나 합성이 꺼져 있으면 바로 돌아오므로, 그때는 잠깐 쉬어 바쁜 대기를 막는다.
                if (!composed || Ms(waitStart, Stopwatch.GetTimestamp()) < 2)
                {
                    Thread.Sleep(8);
                }
                // 아직 그리지 못한 프레임이 있으면 쌓지 않는다.
                if (Interlocked.Exchange(ref _stepQueued, 1) == 0)
                {
                    context.Post(_ =>
                    {
                        Interlocked.Exchange(ref _stepQueued, 0);
                        if (_pumpRunning)
                        {
                            Step();
                        }
                    }, null);
                }
            }
        })
        {
            IsBackground = true,
            Name = "BluePage smooth scroll"
        };
        pump.Start();
    }

    private void StopTimer()
    {
        _timer.Stop();
        _pumpRunning = false;
    }

    public void Dispose()
    {
        StopTimer();
        _timer.Dispose();
        _precaptureTimer.Dispose();
        foreach (var state in _panels.Values)
        {
            DropPrecaptured(state);
        }
    }

    internal static double Ms(long from, long to) => (to - from) * 1000.0 / Stopwatch.Frequency;

    private const uint RDW_INVALIDATE = 0x0001;
    private const uint RDW_ERASE = 0x0004;
    private const uint RDW_ALLCHILDREN = 0x0080;
    private const uint RDW_UPDATENOW = 0x0100;

    [DllImport("user32.dll")]
    internal static extern IntPtr WindowFromPoint(Point point);

    [DllImport("user32.dll")]
    private static extern bool RedrawWindow(IntPtr hWnd, IntPtr updateRect, IntPtr updateRegion, uint flags);

    [DllImport("dwmapi.dll")]
    private static extern int DwmFlush();
}

/// <summary>
/// 스크롤 화면 안의 내용 전체를 한 장으로 찍은 것. 화면에 바로 복사할 수 있는 GDI 비트맵으로 들고 있어서
/// 덮개를 만들 때 변환 작업이 없다. 찍을 때의 스크롤 위치와 화면 크기를 함께 기억한다.
/// 덮개가 가져가면(<see cref="TransferToOverlay"/>) 정리 책임도 덮개로 넘어간다.
/// </summary>
internal sealed class CapturedContent : IDisposable
{
    private IntPtr _bitmap;
    private IntPtr _memoryDc;
    private IntPtr _previousObject;

    public Size ContentSize { get; }
    /// <summary>스크롤 위치가 0일 때의 내용 위치.</summary>
    public Point ContentAtTop { get; }
    public int Scroll { get; }
    public Size ClientSize { get; }
    public long Ticks { get; }

    private CapturedContent(Bitmap image, Point contentAtTop, int scroll, Size clientSize)
    {
        ContentSize = image.Size;
        ContentAtTop = contentAtTop;
        Scroll = scroll;
        ClientSize = clientSize;
        Ticks = Stopwatch.GetTimestamp();

        _bitmap = image.GetHbitmap();
        _memoryDc = CreateCompatibleDC(IntPtr.Zero);
        _previousObject = SelectObject(_memoryDc, _bitmap);
    }

    /// <summary>패널 안의 내용이 하나(PageStack의 내용 묶음)일 때만 찍는다. 실패하면 null.</summary>
    public static CapturedContent? Capture(ScrollableControl panel, int scroll)
    {
        try
        {
            if (panel.Controls.Count != 1)
            {
                return null;
            }
            var content = panel.Controls[0];
            if (content.Width <= 0 || content.Height <= 0)
            {
                return null;
            }

            using var image = new Bitmap(content.Width, content.Height);
            content.DrawToBitmap(image, new Rectangle(Point.Empty, content.Size));
            return new CapturedContent(image, new Point(content.Left, content.Top + scroll), scroll, panel.ClientSize);
        }
        catch (Exception ex) when (ex is ArgumentException or OutOfMemoryException or InvalidOperationException or ExternalException)
        {
            return null;
        }
    }

    /// <summary>지금 화면에 그대로 쓸 수 있는지(스크롤 위치와 크기가 같고 너무 오래되지 않았는지).</summary>
    public bool IsUsable(ScrollableControl panel, int scroll, double maxAgeMs) =>
        _memoryDc != IntPtr.Zero &&
        Scroll == scroll &&
        ClientSize == panel.ClientSize &&
        SmoothScroller.Ms(Ticks, Stopwatch.GetTimestamp()) <= maxAgeMs;

    /// <summary>덮개에 메모리 DC를 넘긴다. 이후 정리는 덮개가 한다.</summary>
    public (IntPtr Bitmap, IntPtr MemoryDc, IntPtr PreviousObject) TransferToOverlay()
    {
        var handles = (_bitmap, _memoryDc, _previousObject);
        _bitmap = IntPtr.Zero;
        _memoryDc = IntPtr.Zero;
        _previousObject = IntPtr.Zero;
        return handles;
    }

    public void Dispose() => Release(_bitmap, _memoryDc, _previousObject);

    internal static void Release(IntPtr bitmap, IntPtr memoryDc, IntPtr previousObject)
    {
        if (memoryDc != IntPtr.Zero)
        {
            SelectObject(memoryDc, previousObject);
            DeleteDC(memoryDc);
        }
        if (bitmap != IntPtr.Zero)
        {
            DeleteObject(bitmap);
        }
    }

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
/// 움직이는 동안 스크롤 패널의 보이는 영역(스크롤바 제외)을 덮는 이미지.
/// 미리 찍어 둔 GDI 비트맵에서 스크롤 위치에 맞춰 보이는 부분만 복사(BitBlt)한다.
/// 이 위를 클릭하면 바로 멈추고 그 자리의 실제 컨트롤에 클릭을 넘긴다.
/// </summary>
internal sealed class SnapshotOverlay : Control
{
    private const int SB_VERT = 1;
    private const int WM_LBUTTONDOWN = 0x0201;
    private const int MK_LBUTTON = 0x0001;
    private const int SRCCOPY = 0x00CC0020;

    private readonly IntPtr _bitmap;
    private readonly IntPtr _memoryDc;
    private readonly IntPtr _previousObject;
    private readonly Size _contentSize;
    private readonly Point _contentAtTop;
    private readonly Action _finish;
    private int _offset;
    private bool _released;

    public ScrollableControl Owner { get; }

    private SnapshotOverlay(ScrollableControl owner, CapturedContent captured, int offset, Action finish)
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.Opaque, true);
        Owner = owner;
        _contentSize = captured.ContentSize;
        _contentAtTop = captured.ContentAtTop;
        _offset = offset;
        _finish = finish;
        TabStop = false;
        (_bitmap, _memoryDc, _previousObject) = captured.TransferToOverlay();
    }

    /// <summary>찍어 둔 이미지로 덮개를 만든다. 이미지는 덮개가 가져가서 정리한다. 실패하면 null.</summary>
    public static SnapshotOverlay? TryCreate(ScrollableControl panel, CapturedContent captured, int scroll, Action finish)
    {
        try
        {
            if (panel.Parent is not { } parent || panel.ClientSize.Width <= 0 || panel.ClientSize.Height <= 0)
            {
                captured.Dispose();
                return null;
            }

            var overlay = new SnapshotOverlay(panel, captured, scroll, finish)
            {
                BackColor = panel.BackColor,
                Bounds = new Rectangle(panel.Left, panel.Top, panel.ClientSize.Width, panel.ClientSize.Height)
            };
            parent.Controls.Add(overlay);
            // 패널 바로 앞에 둔다. 그래야 흐림 띠처럼 패널 앞에 있던 컨트롤이 계속 이미지보다 앞에 보인다.
            parent.Controls.SetChildIndex(overlay, parent.Controls.GetChildIndex(panel));
            overlay.Update();
            return overlay;
        }
        catch (Exception ex) when (ex is ArgumentException or OutOfMemoryException or InvalidOperationException or ExternalException)
        {
            captured.Dispose();
            return null;
        }
    }

    public void ScrollTo(float position)
    {
        var offset = (int)Math.Round(position);
        if (offset == _offset)
        {
            return;
        }
        _offset = offset;
        Invalidate();
        Update();
        // 스크롤바 막대도 함께 움직여 보이게 한다(내용은 움직이지 않고 막대 위치만 바꾼다).
        SetScrollPos(Owner.Handle, SB_VERT, offset, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        var contentRect = new Rectangle(_contentAtTop.X, _contentAtTop.Y - _offset, _contentSize.Width, _contentSize.Height);

        // 내용 바깥(좌우 여백 등)은 배경색으로 채우고, 보이는 내용 부분만 비트맵에서 복사한다.
        using (var background = new SolidBrush(BackColor))
        using (var outside = new Region(ClientRectangle))
        {
            outside.Exclude(contentRect);
            g.FillRegion(background, outside);
        }

        var visible = Rectangle.Intersect(contentRect, ClientRectangle);
        if (visible.Width <= 0 || visible.Height <= 0 || _memoryDc == IntPtr.Zero)
        {
            return;
        }

        var hdc = g.GetHdc();
        try
        {
            BitBlt(hdc, visible.X, visible.Y, visible.Width, visible.Height,
                _memoryDc, visible.X - contentRect.X, visible.Y - contentRect.Y, SRCCOPY);
        }
        finally
        {
            g.ReleaseHdc(hdc);
        }
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        var left = e.Button == MouseButtons.Left;

        // 자기 메시지를 처리하는 도중에 스스로를 지우지 않도록, 마무리는 다음 메시지 차례로 미룬다.
        Owner.BeginInvoke(() =>
        {
            _finish();

            // 이미지가 치워진 뒤 그 자리에 있는 실제 컨트롤에 누름을 넘긴다. 떼는 동작은 그대로 그 컨트롤로 간다.
            if (!left)
            {
                return;
            }
            var screen = Cursor.Position;
            var hwnd = SmoothScroller.WindowFromPoint(screen);
            if (hwnd != IntPtr.Zero)
            {
                var client = screen;
                ScreenToClient(hwnd, ref client);
                PostMessage(hwnd, WM_LBUTTONDOWN, (IntPtr)MK_LBUTTON, (IntPtr)((client.Y << 16) | (client.X & 0xFFFF)));
            }
        });
    }

    protected override void Dispose(bool disposing)
    {
        if (!_released)
        {
            _released = true;
            CapturedContent.Release(_bitmap, _memoryDc, _previousObject);
        }
        base.Dispose(disposing);
    }

    [DllImport("user32.dll")]
    private static extern int SetScrollPos(IntPtr hWnd, int bar, int position, bool redraw);

    [DllImport("user32.dll")]
    private static extern bool ScreenToClient(IntPtr hWnd, ref Point point);

    [DllImport("user32.dll")]
    private static extern IntPtr PostMessage(IntPtr hWnd, int message, IntPtr wParam, IntPtr lParam);

    [DllImport("gdi32.dll")]
    private static extern bool BitBlt(IntPtr hdcDest, int x, int y, int width, int height, IntPtr hdcSrc, int srcX, int srcY, int rop);
}
