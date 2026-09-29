using System.Runtime.InteropServices;

namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>
/// 마우스 휠 스크롤을 부드럽게 만든다. 휠 한 칸마다 목표 위치만 옮기고, 실제 위치는 짧게 감속하며 따라간다.
/// 휠을 연달아 돌리면 목표가 누적돼서 끊기지 않고 이어진다.
/// 등록한 스크롤 패널과 앱 안의 DataGridView(동기화 검토 표)에 적용한다. 표는 줄 단위로만 움직일 수 있어서
/// 한 칸에 한 줄씩 넘긴다. 드롭다운 메뉴처럼 다른 창 위에서 돌린 휠은 건드리지 않는다.
/// </summary>
internal sealed class SmoothScroller : IMessageFilter, IDisposable
{
    private const int WM_MOUSEWHEEL = 0x020A;
    private const float FollowRate = 0.28F;

    private sealed class ScrollState
    {
        public float Current;
        public float Target;
    }

    private readonly Dictionary<ScrollableControl, ScrollState> _panels = new();
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 10 };

    public SmoothScroller()
    {
        _timer.Tick += (_, _) => Step();
    }

    /// <summary>휠 한 칸에 움직이는 거리(96 DPI 기준 픽셀).</summary>
    public int PixelsPerNotch { get; set; } = UiDraw.S(90);

    public void Register(ScrollableControl panel)
    {
        _panels[panel] = new ScrollState();
        panel.Disposed += (_, _) => _panels.Remove(panel);
    }

    public bool PreFilterMessage(ref Message m)
    {
        if (m.Msg != WM_MOUSEWHEEL)
        {
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
        var max = Math.Max(0, panel.DisplayRectangle.Height - panel.ClientSize.Height);
        if (max == 0)
        {
            return false;
        }

        // 스크롤바를 직접 끌었을 수 있으므로, 멈춰 있을 때는 지금 위치에서 다시 시작한다.
        if (!_timer.Enabled)
        {
            state.Current = -panel.AutoScrollPosition.Y;
            state.Target = state.Current;
        }

        state.Target = Math.Clamp(state.Target - delta / 120F * PixelsPerNotch, 0, max);
        _timer.Start();
        return true;
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

    private void Step()
    {
        var moving = false;
        foreach (var (panel, state) in _panels)
        {
            if (panel.IsDisposed)
            {
                continue;
            }
            var remaining = state.Target - state.Current;
            if (Math.Abs(remaining) < 0.5F)
            {
                if (state.Current != state.Target)
                {
                    state.Current = state.Target;
                    panel.AutoScrollPosition = new Point(0, (int)Math.Round(state.Current));
                }
                continue;
            }

            state.Current += remaining * FollowRate;
            panel.AutoScrollPosition = new Point(0, (int)Math.Round(state.Current));
            moving = true;
        }

        if (!moving)
        {
            _timer.Stop();
        }
    }

    public void Dispose() => _timer.Dispose();

    [DllImport("user32.dll")]
    private static extern IntPtr WindowFromPoint(Point point);
}
