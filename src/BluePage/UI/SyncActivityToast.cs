using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using Microsoft365OfficeWebLauncher.Core;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>
/// Windows 11 알림 토스트처럼 화면 우하단(트레이 바로 위)에 떠서 동기화 중인 파일 목록을
/// 창 1개로 관리하는 무포커스 알림. 백그라운드 폴링과 "동기화 검토" 적용에서만 사용되며,
/// 문서 더블클릭으로 여는 헤드리스 프로세스는 이 창을 전혀 쓰지 않는다(NullSyncActivityReporter).
/// 라이트/다크 모드는 AppTheme을 그대로 따라간다.
/// </summary>
public sealed class SyncActivityToast : Form, ISyncActivityReporter
{
    private enum RowState { InProgress, Success, Failure }

    private sealed class ToastRow
    {
        public required StatusDot Icon { get; init; }
        public required Label Name { get; init; }
        public required Label Status { get; init; }
        public RowState State { get; set; }
    }

    private const int MarginFromEdge = 16;
    private const int HideDelayMs = 3000;
    private const double FadeStep = 0.15;

    private readonly Panel _border;
    private readonly TableLayoutPanel _root;
    private readonly Label _appNameLabel;
    private readonly FlowLayoutPanel _itemsPanel;
    private readonly Dictionary<string, ToastRow> _itemRows = new(StringComparer.OrdinalIgnoreCase);
    private readonly System.Windows.Forms.Timer _hideTimer;
    private readonly System.Windows.Forms.Timer _fadeTimer;
    private int _activeCount;
    private double _targetOpacity;
    private bool _notificationsEnabled = true;

    /// <summary>꺼져 있으면 동기화는 그대로 동작하되 이 토스트만 뜨지 않는다(설정 화면의 체크박스와 연결).</summary>
    public bool NotificationsEnabled
    {
        get => _notificationsEnabled;
        set
        {
            _notificationsEnabled = value;
            if (!value && Visible)
            {
                _hideTimer.Stop();
                _fadeTimer.Stop();
                Hide();
                ClearItems();
            }
        }
    }

    public SyncActivityToast()
    {
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        ShowInTaskbar = false;
        TopMost = true;
        AllowTransparency = true;
        Opacity = 0;
        Font = new Font("Segoe UI", 9F);
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        MinimumSize = new Size(260, 0);
        MaximumSize = new Size(340, 300);

        _border = new Panel { Padding = new Padding(1), AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
        _root = new TableLayoutPanel
        {
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(14, 12, 14, 12)
        };
        _border.Controls.Add(_root);
        Controls.Add(_border);

        _appNameLabel = new Label
        {
            Text = AppBrand.Name,
            AutoSize = true,
            Font = new Font(Font.FontFamily, 8F),
            Margin = new Padding(0, 0, 0, 8)
        };
        _root.Controls.Add(_appNameLabel);

        _itemsPanel = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0)
        };
        _root.Controls.Add(_itemsPanel);

        _hideTimer = new System.Windows.Forms.Timer { Interval = HideDelayMs };
        _hideTimer.Tick += (_, _) =>
        {
            _hideTimer.Stop();
            BeginFade(0.0);
        };

        _fadeTimer = new System.Windows.Forms.Timer { Interval = 15 };
        _fadeTimer.Tick += (_, _) => StepFade();

        RefreshTheme();
        AppTheme.Changed += RefreshTheme;
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        AppTheme.ApplyTitleBarTheme(Handle, AppTheme.IsDark);
    }

    /// <summary>테마가 바뀔 때(수동 전환 또는 Windows 시스템 테마 변경) 이미 떠 있는 항목까지 다시 칠한다.</summary>
    private void RefreshTheme()
    {
        var theme = AppTheme.Current;

        BackColor = theme.CardBackground;
        _border.BackColor = theme.Border;
        _root.BackColor = theme.CardBackground;
        _appNameLabel.ForeColor = theme.TextSecondary;

        foreach (var row in _itemRows.Values)
        {
            row.Name.ForeColor = theme.TextPrimary;
            ApplyRowState(row, row.State, theme);
        }

        if (IsHandleCreated)
        {
            AppTheme.ApplyTitleBarTheme(Handle, AppTheme.IsDark);
        }
    }

    // 포커스를 뺏지 않고, 네이티브 그림자가 붙는 알림 창(Windows 토스트와 동일한 방식)
    protected override bool ShowWithoutActivation => true;

    protected override CreateParams CreateParams
    {
        get
        {
            const int WS_EX_NOACTIVATE = 0x08000000;
            const int CS_DROPSHADOW = 0x00020000;
            var cp = base.CreateParams;
            cp.ExStyle |= WS_EX_NOACTIVATE;
            cp.ClassStyle |= CS_DROPSHADOW;
            return cp;
        }
    }

    public void ReportStarted(string localFilePath)
    {
        _hideTimer.Stop();
        _activeCount++;

        var fileName = Path.GetFileName(localFilePath);
        if (!_itemRows.TryGetValue(localFilePath, out var row))
        {
            _itemsPanel.Controls.Add(BuildRow(fileName, out row));
            _itemRows[localFilePath] = row;
        }

        SetRowState(row, RowState.InProgress);
        PositionAndShow();
    }

    public void ReportCompleted(string localFilePath, bool success)
    {
        _activeCount = Math.Max(0, _activeCount - 1);

        if (_itemRows.TryGetValue(localFilePath, out var row))
        {
            SetRowState(row, success ? RowState.Success : RowState.Failure);
        }

        if (_activeCount == 0)
        {
            _hideTimer.Stop();
            _hideTimer.Start();
        }
    }

    private void SetRowState(ToastRow row, RowState state)
    {
        row.State = state;
        ApplyRowState(row, state, AppTheme.Current);
    }

    private static void ApplyRowState(ToastRow row, RowState state, ThemePalette theme)
    {
        switch (state)
        {
            case RowState.InProgress:
                row.Icon.Style = StatusDot.DotStyle.Filled;
                row.Icon.DotColor = theme.Accent;
                row.Status.Text = "동기화 중…";
                row.Status.ForeColor = theme.Accent;
                break;
            case RowState.Success:
                row.Icon.Style = StatusDot.DotStyle.Check;
                row.Icon.DotColor = theme.Success;
                row.Status.Text = "동기화 완료";
                row.Status.ForeColor = theme.Success;
                break;
            case RowState.Failure:
                row.Icon.Style = StatusDot.DotStyle.Cross;
                row.Icon.DotColor = theme.Failure;
                row.Status.Text = "실패";
                row.Status.ForeColor = theme.Failure;
                break;
        }
    }

    /// <summary>파일명 + 상태 아이콘 한 줄(2행: 파일명 / 작은 상태 텍스트)을 만든다.</summary>
    private Control BuildRow(string fileName, out ToastRow row)
    {
        var theme = AppTheme.Current;
        var table = new TableLayoutPanel
        {
            ColumnCount = 2,
            RowCount = 2,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0, 4, 0, 4)
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var icon = new StatusDot { Margin = new Padding(0, 3, 8, 0) };
        table.Controls.Add(icon, 0, 0);
        table.SetRowSpan(icon, 2);

        var name = new Label
        {
            Text = fileName,
            AutoSize = true,
            MaximumSize = new Size(220, 0),
            ForeColor = theme.TextPrimary,
            Margin = new Padding(0)
        };
        table.Controls.Add(name, 1, 0);

        var status = new Label
        {
            AutoSize = true,
            Font = new Font(Font.FontFamily, 8F),
            Margin = new Padding(0, 1, 0, 0)
        };
        table.Controls.Add(status, 1, 1);

        row = new ToastRow { Icon = icon, Name = name, Status = status };
        return table;
    }

    private void PositionAndShow()
    {
        if (!NotificationsEnabled)
        {
            return;
        }

        var wasVisible = Visible;
        if (!wasVisible)
        {
            Opacity = 0;
            Show();
        }

        // FlowLayoutPanel→TableLayoutPanel→Panel→Form으로 중첩된 AutoSize 컨테이너는 항목이
        // 추가된 직후 Width/Height 재계산이 한 박자 늦게 반영될 때가 있어(다음 레이아웃 패스까지 지연),
        // 위치를 잡기 전에 레이아웃을 강제로 갱신하고 PreferredSize로 크기를 직접 맞춘다.
        PerformLayout();
        Size = PreferredSize;

        var workingArea = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
        Location = new Point(workingArea.Right - Width - MarginFromEdge, workingArea.Bottom - Height - MarginFromEdge);

        BeginFade(1.0);
    }

    private void BeginFade(double target)
    {
        _targetOpacity = target;
        if (!_fadeTimer.Enabled)
        {
            _fadeTimer.Start();
        }
    }

    private void StepFade()
    {
        if (Opacity < _targetOpacity)
        {
            Opacity = Math.Min(_targetOpacity, Opacity + FadeStep);
        }
        else if (Opacity > _targetOpacity)
        {
            Opacity = Math.Max(_targetOpacity, Opacity - FadeStep);
        }

        if (Opacity != _targetOpacity)
        {
            return;
        }

        _fadeTimer.Stop();
        if (_targetOpacity <= 0.0)
        {
            Hide();
            ClearItems();
        }
    }

    private void ClearItems()
    {
        _itemsPanel.Controls.Clear();
        _itemRows.Clear();
        _activeCount = 0;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            AppTheme.Changed -= RefreshTheme;
        }

        base.Dispose(disposing);
    }

    /// <summary>동기화 중(●)/완료(✓)/실패(✕) 상태를 부드럽게 그려주는 작은 아이콘. 폰트 글리프 대신 직접 그려 더 또렷하다.</summary>
    private sealed class StatusDot : Control
    {
        public enum DotStyle { Filled, Check, Cross }

        private Color _dotColor = Color.Gray;
        private DotStyle _style = DotStyle.Filled;

        public Color DotColor
        {
            get => _dotColor;
            set { _dotColor = value; Invalidate(); }
        }

        public DotStyle Style
        {
            get => _style;
            set { _style = value; Invalidate(); }
        }

        public StatusDot()
        {
            SetStyle(
                ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.SupportsTransparentBackColor,
                true);
            BackColor = Color.Transparent;
            Size = new Size(16, 16);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var rect = new Rectangle(1, 1, Width - 3, Height - 3);

            switch (Style)
            {
                case DotStyle.Filled:
                    using (var brush = new SolidBrush(DotColor))
                    {
                        e.Graphics.FillEllipse(brush, new Rectangle(rect.X + 3, rect.Y + 3, rect.Width - 6, rect.Height - 6));
                    }
                    break;

                case DotStyle.Check:
                    using (var pen = new Pen(DotColor, 2f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                    {
                        e.Graphics.DrawLines(pen, new[]
                        {
                            new PointF(rect.Left + 1, rect.Top + rect.Height * 0.55f),
                            new PointF(rect.Left + rect.Width * 0.4f, rect.Bottom - 2),
                            new PointF(rect.Right - 1, rect.Top + 1)
                        });
                    }
                    break;

                case DotStyle.Cross:
                    using (var pen = new Pen(DotColor, 2f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                    {
                        e.Graphics.DrawLine(pen, rect.Left + 1, rect.Top + 1, rect.Right - 1, rect.Bottom - 1);
                        e.Graphics.DrawLine(pen, rect.Right - 1, rect.Top + 1, rect.Left + 1, rect.Bottom - 1);
                    }
                    break;
            }
        }
    }
}
