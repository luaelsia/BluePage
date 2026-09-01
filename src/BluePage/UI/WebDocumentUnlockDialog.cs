using Microsoft365OfficeWebLauncher.OneDrive;
using Microsoft365OfficeWebLauncher.Core;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>화면 우하단에서 웹 잠금과 로컬 파일 안정 상태를 자동 확인하고 사용자의 동기화 결정을 기다린다.</summary>
public sealed class WebDocumentUnlockDialog : Form
{
    private const int MarginFromEdge = 16;
    private const int StatusCheckIntervalMs = 2000;
    private const int RequiredStableChecks = 3;
    private static readonly TimeSpan ReadyTimeout = TimeSpan.FromSeconds(60);

    private readonly string _localFilePath;
    private readonly Func<CancellationToken, Task<RemoteLockState>> _checkRemoteLock;
    private readonly Func<CancellationToken, Task<SyncResult>> _retrySync;
    private readonly CancellationTokenSource _cancellation = new();
    private readonly Panel _border;
    private readonly TableLayoutPanel _layout;
    private readonly Label _statusLabel;
    private readonly ProgressBar _progressBar;
    private readonly Button _syncButton;
    private readonly System.Windows.Forms.Timer _statusTimer;

    private LocalFileState? _lastLocalState;
    private int _stableCheckCount;
    private DateTimeOffset? _readySinceUtc;
    private bool _checkInProgress;
    private bool _finished;

    public SyncResult? SyncResult { get; private set; }
    public Exception? SyncError { get; private set; }
    public bool SynchronizationStopped { get; private set; }
    public bool TargetUnavailable { get; private set; }

    public WebDocumentUnlockDialog(
        string localFilePath,
        Func<CancellationToken, Task<RemoteLockState>> checkRemoteLock,
        Func<CancellationToken, Task<SyncResult>> retrySync)
    {
        _localFilePath = localFilePath;
        _checkRemoteLock = checkRemoteLock;
        _retrySync = retrySync;

        Text = "동기화 준비";
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        ShowInTaskbar = false;
        TopMost = true;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        MinimumSize = new Size(360, 0);
        MaximumSize = new Size(400, 340);
        Font = new Font("Segoe UI", 9F);

        _border = new Panel { Padding = new Padding(1), AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
        _layout = new TableLayoutPanel
        {
            Padding = new Padding(16, 13, 16, 14),
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };
        _border.Controls.Add(_layout);
        Controls.Add(_border);

        _layout.Controls.Add(new Label
        {
            Text = $"{AppBrand.Name}  ·  동기화 준비",
            AutoSize = true,
            Font = new Font(Font.FontFamily, 8F, FontStyle.Bold),
            Tag = ThemeApplier.SecondaryTag,
            Margin = new Padding(0, 0, 0, 9)
        });

        _layout.Controls.Add(new Label
        {
            Text = $"'{Path.GetFileName(localFilePath)}' 문서가 웹에서 열려 있어 지금은 동기화할 수 없습니다.\n\n" +
                   "웹 문서를 닫고 로컬 파일 수정을 완료해 주세요.\n" +
                   "준비되면 지금 동기화하거나, 나중에 문서를 다시 열 때 동기화할 수 있습니다.",
            AutoSize = true,
            MaximumSize = new Size(350, 0),
            Margin = new Padding(0, 0, 0, 12)
        });

        _progressBar = new ProgressBar
        {
            Style = ProgressBarStyle.Marquee,
            MarqueeAnimationSpeed = 30,
            Width = 350,
            Height = 7,
            Margin = new Padding(0, 0, 0, 10)
        };
        _layout.Controls.Add(_progressBar);

        _statusLabel = new Label
        {
            Text = "웹 잠금과 로컬 파일 상태를 확인하는 중…",
            AutoSize = true,
            MaximumSize = new Size(350, 0),
            Tag = ThemeApplier.SecondaryTag,
            Margin = new Padding(0, 0, 0, 14)
        };
        _layout.Controls.Add(_statusLabel);

        var buttonPanel = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Anchor = AnchorStyles.Right,
            Margin = new Padding(0)
        };

        _syncButton = new Button { Text = "동기화", AutoSize = true, Enabled = false };
        _syncButton.Click += async (_, _) => await StartSyncAsync();
        var stopButton = new Button { Text = "나중에", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
        stopButton.Click += (_, _) => StopSynchronization("사용자가 나중에 동기화하기로 했습니다.");
        buttonPanel.Controls.Add(_syncButton);
        buttonPanel.Controls.Add(stopButton);
        AcceptButton = _syncButton;
        CancelButton = stopButton;
        _layout.Controls.Add(buttonPanel);

        _statusTimer = new System.Windows.Forms.Timer { Interval = StatusCheckIntervalMs };
        _statusTimer.Tick += async (_, _) => await CheckStatusAsync();
        Shown += async (_, _) =>
        {
            PositionAtBottomRight();
            _statusTimer.Start();
            await CheckStatusAsync();
        };
        FormClosing += (_, _) =>
        {
            if (!_finished)
            {
                SynchronizationStopped = true;
                _cancellation.Cancel();
            }
        };

        AppTheme.Changed += RefreshTheme;
        RefreshTheme();
    }

    private async Task CheckStatusAsync()
    {
        if (_finished || _checkInProgress) return;

        _checkInProgress = true;
        try
        {
            if (!LocalFileStateMonitor.TryRead(_localFilePath, out var currentState))
            {
                TargetUnavailable = true;
                StopSynchronization("로컬 파일이 이동되었거나 삭제되어 알림을 닫았습니다.", deferSynchronization: false);
                return;
            }
            if (_lastLocalState == currentState) _stableCheckCount++;
            else
            {
                _lastLocalState = currentState;
                _stableCheckCount = 0;
            }

            var localStable = _stableCheckCount >= RequiredStableChecks;
            var lockState = await _checkRemoteLock(_cancellation.Token);
            if (!localStable || lockState == RemoteLockState.Locked)
            {
                SetWaitingState(localStable, lockState);
                return;
            }

            if (lockState == RemoteLockState.Unlocked)
            {
                _readySinceUtc ??= DateTimeOffset.UtcNow;
                if (DateTimeOffset.UtcNow - _readySinceUtc.Value >= ReadyTimeout)
                {
                    StopSynchronization("동기화 가능 상태가 60초 동안 유지되어 알림을 자동으로 닫았습니다.");
                    return;
                }
            }
            else _readySinceUtc = null;

            _syncButton.Enabled = true;
            _progressBar.Style = ProgressBarStyle.Blocks;
            _progressBar.MarqueeAnimationSpeed = 0;
            var lockNote = lockState == RemoteLockState.Unknown ? " (웹 잠금은 동기화 시 다시 확인)" : string.Empty;
            _statusLabel.Text = $"동기화 가능{lockNote} · 지금 동기화하거나 나중에 다시 열어도 됩니다.";
        }
        catch (OperationCanceledException) when (_cancellation.IsCancellationRequested)
        {
        }
        catch (Exception ex)
        {
            _readySinceUtc = null;
            _syncButton.Enabled = false;
            _statusLabel.Text = $"상태 확인에 실패했습니다. 자동으로 다시 확인합니다. ({ex.Message})";
        }
        finally
        {
            _checkInProgress = false;
        }
    }

    private void SetWaitingState(bool localStable, RemoteLockState lockState)
    {
        _readySinceUtc = null;
        _syncButton.Enabled = false;
        _progressBar.Style = ProgressBarStyle.Marquee;
        _progressBar.MarqueeAnimationSpeed = 30;
        _statusLabel.Text = !localStable
            ? "로컬 파일 수정이 끝나기를 기다리는 중…"
            : lockState == RemoteLockState.Locked
                ? "웹 문서가 닫히기를 기다리는 중…"
                : "동기화 상태를 확인하는 중…";
    }

    private async Task StartSyncAsync()
    {
        _statusTimer.Stop();
        _syncButton.Enabled = false;
        _progressBar.Style = ProgressBarStyle.Marquee;
        _progressBar.MarqueeAnimationSpeed = 30;
        _statusLabel.Text = "동기화 중…";
        try
        {
            SyncResult = await _retrySync(_cancellation.Token);
            _finished = true;
            DialogResult = DialogResult.OK;
            Close();
        }
        catch (Exception ex) when (GraphErrorHelper.IsResourceLocked(ex))
        {
            _readySinceUtc = null;
            _stableCheckCount = 0;
            _statusLabel.Text = "웹 문서가 아직 열려 있습니다. 자동 확인을 계속합니다.";
            _statusTimer.Start();
        }
        catch (OperationCanceledException) when (_cancellation.IsCancellationRequested)
        {
        }
        catch (Exception ex)
        {
            SyncError = ex;
            _finished = true;
            DialogResult = DialogResult.Abort;
            Close();
        }
    }

    private void StopSynchronization(string reason, bool deferSynchronization = true)
    {
        if (_finished) return;
        SynchronizationStopped = deferSynchronization;
        _finished = true;
        _statusTimer.Stop();
        _statusLabel.Text = reason;
        DialogResult = DialogResult.Cancel;
        Close();
    }

    private void PositionAtBottomRight()
    {
        PerformLayout();
        Size = PreferredSize;
        var workingArea = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
        Location = new Point(workingArea.Right - Width - MarginFromEdge, workingArea.Bottom - Height - MarginFromEdge);
    }

    private void RefreshTheme()
    {
        ThemeApplier.Apply(this, AppTheme.Current);
        _border.BackColor = AppTheme.Current.Border;
        _layout.BackColor = AppTheme.Current.CardBackground;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            AppTheme.Changed -= RefreshTheme;
            _statusTimer.Dispose();
            _cancellation.Dispose();
        }
        base.Dispose(disposing);
    }
}
