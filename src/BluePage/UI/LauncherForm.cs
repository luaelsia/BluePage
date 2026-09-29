using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Threading;
using Microsoft365OfficeWebLauncher.Auth;
using Microsoft365OfficeWebLauncher.Config;
using Microsoft365OfficeWebLauncher.Core;
using Microsoft365OfficeWebLauncher.Cloud;
using Microsoft365OfficeWebLauncher.GoogleDrive;
using Microsoft365OfficeWebLauncher.Logging;
using Microsoft365OfficeWebLauncher.OneDrive;
using Microsoft365OfficeWebLauncher.Registry;
using Microsoft365OfficeWebLauncher.Update;
using Microsoft365OfficeWebLauncher.UI.Controls;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>
/// 설정 창. 왼쪽 사이드 메뉴로 홈 / 문서 연결 / 설정 화면을 오간다.
/// 이 파일은 동작(로그인, 동기화, 업데이트, 트레이)을 맡고, 화면 구성은 LauncherForm.*.cs 에 나눠 둔다.
/// 인수 없이 실행하거나 --settings/--minimized로 실행했을 때만 뜨며, 문서를 여는 헤드리스 흐름에는 관여하지 않는다.
/// </summary>
public sealed partial class LauncherForm : Form
{
    private readonly AppConfig _config;
    private readonly FileLogger _logger;
    private readonly GraphAuthService _authService;
    private readonly GoogleAuthService _googleAuthService;
    private readonly OneDriveUploadService _uploadService;
    private readonly GoogleDriveService _googleDriveService;
    private readonly UploadManifest _manifest;
    private readonly LaunchOrchestrator _orchestrator;
    private readonly FileAssociationRegistrar _registrar;
    private readonly StartupRegistrar _startupRegistrar;
    private readonly string _exePath;
    private readonly bool _launchMinimizedToTray;

    private NotifyIcon _trayIcon = null!;
    private System.Windows.Forms.Timer _backgroundSyncTimer = null!;
    private EventWaitHandle _showWindowEvent = null!;
    private RegisteredWaitHandle? _showWindowRegisteredWait;
    private SyncActivityToast _activityToast = null!;
    private ToolStripMenuItem _trayUpdateItem = null!;
    private ToolStripSeparator _trayUpdateSeparator = null!;
    private System.Windows.Forms.Timer _updateTimer = null!;
    private UpdateInfo? _availableUpdate;
    private string? _balloonShownForVersion;
    private bool _updateDialogOpen;

    private const int MinSyncIntervalSeconds = 1;
    private const int MaxSyncIntervalSeconds = 600;

    private bool _hasCachedAccount;
    private bool _hasCachedGoogleAccount;
    private bool _isExiting;
    private bool _syncInProgress;
    private bool _backgroundSyncRunning;
    private bool _openSyncReviewAfterBackgroundSync;

    public LauncherForm(
        AppConfig config,
        FileLogger logger,
        GraphAuthService authService,
        GoogleAuthService googleAuthService,
        OneDriveUploadService uploadService,
        GoogleDriveService googleDriveService,
        UploadManifest manifest,
        LaunchOrchestrator orchestrator,
        FileAssociationRegistrar registrar,
        StartupRegistrar startupRegistrar,
        string exePath,
        bool launchMinimizedToTray = false)
    {
        _config = config;
        _logger = logger;
        _authService = authService;
        _googleAuthService = googleAuthService;
        _uploadService = uploadService;
        _googleDriveService = googleDriveService;
        _manifest = manifest;
        _orchestrator = orchestrator;
        _registrar = registrar;
        _startupRegistrar = startupRegistrar;
        _exePath = exePath;
        _launchMinimizedToTray = launchMinimizedToTray;

        BuildUi();
        // 마우스 휠 스크롤을 부드럽게: 앱 안의 휠 메시지를 먼저 받아 등록한 화면과 표에 적용한다.
        Application.AddMessageFilter(_smoothScroller);
        BuildTrayIcon();
        BuildActivityToast();
        BuildBackgroundSyncTimer();
        BuildUpdateTimer();
        BuildSingleInstanceListener();

        ApplyTheme();
        AppTheme.Changed += ApplyTheme;
        Microsoft.Win32.SystemEvents.UserPreferenceChanged += OnSystemThemeChanged;

        Load += async (_, _) => await OnLoadAsync();
        Shown += OnShown;
        FormClosing += OnFormClosing;
    }

    private async Task OnLoadAsync()
    {
        EnsureFileAssociationRegistered();
        RefreshFileAssocStatus();
        RefreshSyncStatus();
        LoadAutoStartState();
        await RefreshAccountStatusAsync();
        await RefreshGoogleAccountStatusAsync();
    }

    private void OnShown(object? sender, EventArgs e)
    {
        CreateScrollFades();
        if (_launchMinimizedToTray)
        {
            Hide();
            _trayIcon.Visible = true;
        }
    }

    private void OnThemeChanged()
    {
        var preference = _themeSelector.SelectedIndex switch
        {
            1 => ThemePreference.Light,
            2 => ThemePreference.Dark,
            _ => ThemePreference.System
        };

        AppTheme.SetPreference(preference);
        _config.Theme = preference.ToString();
        ConfigLoader.Save(_config);
        _logger.Info($"테마 설정 변경: {preference}");
    }

    /// <summary>테마 선택 또는 Windows 시스템 테마 변경(AppTheme.Changed) 시 창과 트레이 메뉴를 다시 칠한다.</summary>
    private void ApplyTheme()
    {
        ThemeApplier.Apply(this, AppTheme.Current);

        if (_trayIcon.ContextMenuStrip is { } menu)
        {
            ThemedMenu.Apply(menu);
        }
    }

    private void OnSystemThemeChanged(object? sender, Microsoft.Win32.UserPreferenceChangedEventArgs e) =>
        AppTheme.NotifySystemThemeChanged();

    /// <summary>시작 30초 뒤 한 번, 그 뒤로 24시간마다 새 버전을 확인한다(자동 확인을 끄면 건너뜀).</summary>
    private void BuildUpdateTimer()
    {
        _updateTimer = new System.Windows.Forms.Timer { Interval = 30_000 };
        _updateTimer.Tick += async (_, _) =>
        {
            _updateTimer.Interval = (int)TimeSpan.FromHours(24).TotalMilliseconds;
            await CheckForUpdatesAsync(manual: false);
        };
        _updateTimer.Start();
    }

    /// <summary>
    /// manual이면 [업데이트 확인] 버튼에서 부른 것: 결과를 상태 문구로 보여 주고, 새 버전이면 바로 업데이트 창을 연다.
    /// 자동 확인은 실패해도 조용히 넘어가고, 건너뛰기로 한 버전은 알리지 않는다.
    /// </summary>
    private async Task CheckForUpdatesAsync(bool manual)
    {
        if (!manual && !_config.CheckForUpdates)
        {
            return;
        }

        _checkUpdateButton.Enabled = false;
        if (manual)
        {
            _updateStatusLabel.Text = "확인하는 중...";
        }

        try
        {
            var update = await UpdateChecker.GetNewerReleaseAsync(CancellationToken.None);
            if (update is null)
            {
                _availableUpdate = null;
                _updateStatusLabel.Text = "최신 버전입니다.";
                SetTrayUpdateItem(null);
                return;
            }

            _availableUpdate = update;
            var skipped = string.Equals(_config.SkippedUpdateVersion, update.DisplayVersion, StringComparison.OrdinalIgnoreCase);
            _updateStatusLabel.Text = skipped && !manual
                ? $"새 버전 {update.DisplayVersion} (건너뛰기로 설정함)"
                : $"새 버전 {update.DisplayVersion}을(를) 설치할 수 있습니다.";
            _logger.Info($"새 버전 발견: {update.DisplayVersion} (현재 {AppBrand.Version})");

            if (manual)
            {
                SetTrayUpdateItem(update);
                ShowUpdateDialog();
                return;
            }

            if (skipped)
            {
                SetTrayUpdateItem(null);
                return;
            }

            SetTrayUpdateItem(update);
            if (_balloonShownForVersion != update.DisplayVersion)
            {
                _balloonShownForVersion = update.DisplayVersion;
                _trayIcon.ShowBalloonTip(
                    10_000,
                    $"{AppBrand.Name} 업데이트",
                    $"새 버전 {update.DisplayVersion}을(를) 설치할 수 있습니다. 여기를 누르면 업데이트 창이 열립니다.",
                    ToolTipIcon.Info);
            }
        }
        catch (Exception ex)
        {
            _logger.Warn($"업데이트 확인 실패: {ex.Message}");
            if (manual)
            {
                _updateStatusLabel.Text = "확인하지 못했습니다. 인터넷 연결을 확인해 주세요.";
            }
        }
        finally
        {
            _checkUpdateButton.Enabled = true;
        }
    }

    private void SetTrayUpdateItem(UpdateInfo? update)
    {
        _trayUpdateItem.Visible = update is not null;
        _trayUpdateSeparator.Visible = update is not null;
        if (update is not null)
        {
            _trayUpdateItem.Text = $"업데이트 설치 (v{update.DisplayVersion})…";
        }
    }

    private void ShowUpdateDialog()
    {
        if (_availableUpdate is not { } update || _updateDialogOpen)
        {
            return;
        }

        _updateDialogOpen = true;
        try
        {
            using var dialog = new UpdateDialog(update);
            var result = Visible ? dialog.ShowDialog(this) : dialog.ShowDialog();

            if (result == DialogResult.Ignore)
            {
                _config.SkippedUpdateVersion = update.DisplayVersion;
                ConfigLoader.Save(_config);
                _logger.Info($"업데이트 건너뛰기: {update.DisplayVersion}");
                _updateStatusLabel.Text = $"새 버전 {update.DisplayVersion} (건너뛰기로 설정함)";
                SetTrayUpdateItem(null);
                return;
            }

            if (result == DialogResult.OK && dialog.InstallerPath is { } installerPath)
            {
                _logger.Info($"업데이트 설치 시작: {update.DisplayVersion} ({installerPath})");
                UpdateChecker.LaunchSilentInstall(installerPath);
                // 설치 파일이 실행 중인 Blue Page를 끝내고 설치 후 다시 실행한다. 여기서는 먼저 스스로 종료한다.
                ExitApplication();
            }
        }
        catch (Exception ex)
        {
            _logger.Error("업데이트 설치 시작 실패", ex);
            AppMessageDialog.Show($"업데이트를 시작하지 못했습니다.\n\n{ex.Message}", AppBrand.Name, AppMessageKind.Warning);
        }
        finally
        {
            _updateDialogOpen = false;
        }
    }

    private void BuildTrayIcon()
    {
        var menu = new ContextMenuStrip();

        // 트레이에는 창을 열지 않고 바로 쓰는 동작만 둔다. 계정, 파일 연결, 설정 항목은 창에서 다룬다.
        // 업데이트 항목은 새 버전을 찾았을 때만 맨 위에 보인다.
        _trayUpdateItem = new ToolStripMenuItem("업데이트 설치") { Visible = false, Font = new Font(menu.Font, FontStyle.Bold) };
        _trayUpdateItem.Click += (_, _) => ShowUpdateDialog();
        menu.Items.Add(_trayUpdateItem);
        _trayUpdateSeparator = new ToolStripSeparator { Visible = false };
        menu.Items.Add(_trayUpdateSeparator);

        menu.Items.Add("열기", null, (_, _) => ShowFromTray());
        menu.Items.Add(new ToolStripSeparator());

        menu.Items.Add("동기화 검토…", null, async (_, _) => await OnSyncActionAsync());
        menu.Items.Add("OneDrive에서 보기…", null, async (_, _) => await OnOpenOneDriveClickedAsync());
        menu.Items.Add("Google Drive에서 보기…", null, async (_, _) => await OnOpenGoogleDriveClickedAsync());
        menu.Items.Add("백업 폴더 열기", null, (_, _) => OpenBackupFolder());
        menu.Items.Add(new ToolStripSeparator());

        menu.Items.Add("종료", null, (_, _) => ExitApplication());

        _trayIcon = new NotifyIcon
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? SystemIcons.Application,
            Text = AppBrand.Name,
            Visible = true, // 창이 보이는 상태에서도 항상 트레이에 표시(프로그램이 켜져 있는 동안 계속)
            ContextMenuStrip = menu
        };
        _trayIcon.DoubleClick += (_, _) => ShowFromTray();
        // 트레이 풍선은 업데이트 알림에만 쓴다.
        _trayIcon.BalloonTipClicked += (_, _) => ShowUpdateDialog();
    }

    /// <summary>
    /// 백그라운드 폴링·"동기화 검토" 적용에서 실제로 업로드/다운로드가 일어날 때 우하단에 뜨는 토스트를 연결한다.
    /// 더블클릭으로 여는 헤드리스 프로세스는 이 GUI와 별개 프로세스라 애초에 이 코드 경로를 타지 않는다.
    /// </summary>
    private void BuildActivityToast()
    {
        _activityToast = new SyncActivityToast { NotificationsEnabled = _showToastCheckBox.Checked };
        _orchestrator.AttachActivityReporter(_activityToast);
    }

    private void OnShowToastCheckedChanged()
    {
        _activityToast.NotificationsEnabled = _showToastCheckBox.Checked;
        _config.ShowSyncToast = _showToastCheckBox.Checked;
        ConfigLoader.Save(_config);
        _logger.Info($"동기화 알림 표시 설정 변경: {_showToastCheckBox.Checked}");
    }

    private void BuildBackgroundSyncTimer()
    {
        // 창을 열어두지 않아도(트레이 상주 상태에서도) 주기적으로 온라인 변경 사항을 로컬에 반영한다.
        // Graph 변경 알림(webhook)은 공개 HTTPS 엔드포인트가 필요해 로컬 앱에는 쓸 수 없으므로 폴링 방식을 사용한다.
        // 주기는 GUI의 "자동 동기화 주기" 입력으로 조절 가능(config.json의 backgroundSyncIntervalSeconds).
        var seconds = Math.Clamp(_config.BackgroundSyncIntervalSeconds, MinSyncIntervalSeconds, MaxSyncIntervalSeconds);
        _backgroundSyncTimer = new System.Windows.Forms.Timer { Interval = seconds * 1000 };
        _backgroundSyncTimer.Tick += async (_, _) => await OnBackgroundSyncTickAsync();
        _backgroundSyncTimer.Start();
    }

    /// <summary>
    /// Program.cs가 이미 이 GUI 인스턴스가 실행 중임을 감지하면(중복 실행 시도) 새 프로세스를 띄우는 대신
    /// 이 이벤트를 신호해 기존 창을 앞으로 가져오도록 한다.
    /// </summary>
    private void BuildSingleInstanceListener()
    {
        _showWindowEvent = new EventWaitHandle(false, EventResetMode.AutoReset, SingleInstance.ShowWindowEventName);
        _showWindowRegisteredWait = ThreadPool.RegisterWaitForSingleObject(
            _showWindowEvent,
            (_, _) =>
            {
                try
                {
                    if (!IsDisposed)
                    {
                        BeginInvoke(new Action(ShowFromTray));
                    }
                }
                catch (ObjectDisposedException)
                {
                    // 창이 막 닫히는 중이었던 경우 — 무시
                }
            },
            null,
            Timeout.Infinite,
            executeOnlyOnce: false);
    }

    private void ApplySyncInterval(int seconds)
    {
        seconds = Math.Clamp(seconds, MinSyncIntervalSeconds, MaxSyncIntervalSeconds);
        if (_backgroundSyncTimer is null ||
            (seconds == _config.BackgroundSyncIntervalSeconds && _backgroundSyncTimer.Interval == seconds * 1000))
        {
            return;
        }
        _config.BackgroundSyncIntervalSeconds = seconds;
        ConfigLoader.Save(_config);
        _backgroundSyncTimer.Interval = seconds * 1000;
        _logger.Info($"자동 동기화 주기 변경: {seconds}초");
    }

    private async Task OnBackgroundSyncTickAsync()
    {
        if (_syncInProgress || (!_hasCachedAccount && !_hasCachedGoogleAccount))
        {
            return;
        }

        // 더블클릭으로 문서를 열면 매번 별도의 헤드리스 프로세스가 떠서 manifest.json에 직접 기록한다.
        // 이 GUI 프로세스는 그 변경을 자동으로 알 수 없으므로, 동기화 시도 전 항상 디스크에서 다시 읽어온다
        // (안 그러면 방금 다른 프로세스가 이미 처리한 변경을 "충돌"로 잘못 판단해 매번 대화상자가 반복된다).
        _manifest.Reload();

        if (_manifest.AllEntries.Count == 0)
        {
            return;
        }

        _syncInProgress = true;
        _backgroundSyncRunning = true;
        _logger.Debug("백그라운드 자동 동기화 시작");

        try
        {
            await _orchestrator.SyncAllAsync(CancellationToken.None);
        }
        catch (Exception ex)
        {
            _logger.Warn($"백그라운드 자동 동기화 실패: {ex.Message}");
        }
        finally
        {
            _syncInProgress = false;
            _backgroundSyncRunning = false;
            if (!IsDisposed)
            {
                RefreshSyncStatus();
                // 동기화 도중 Google 로그인이 만료됐을 수 있으므로 계정 상태 표시를 같이 갱신한다.
                await RefreshGoogleAccountStatusAsync();

                if (_openSyncReviewAfterBackgroundSync)
                {
                    _openSyncReviewAfterBackgroundSync = false;
                    _syncActionButton.Enabled = true;
                    _syncActionButton.Text = SyncReviewButtonText;
                    await OnSyncActionAsync();
                }
            }
        }
    }

    private void OnFormClosing(object? sender, FormClosingEventArgs e)
    {
        if (_isExiting || e.CloseReason != CloseReason.UserClosing)
        {
            return;
        }

        // 창을 닫아도 앱은 종료하지 않고 백그라운드로 들어간다(트레이 아이콘은 항상 떠 있음). 완전히 끄려면 트레이 메뉴의 "종료"를 사용한다.
        e.Cancel = true;
        Hide();
    }

    private void ShowFromTray()
    {
        Show();
        WindowState = FormWindowState.Normal;
        Activate();
    }

    private void ExitApplication()
    {
        _isExiting = true;
        _trayIcon.Visible = false;
        Close();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _trayIcon?.Dispose();
            _backgroundSyncTimer?.Dispose();
            _updateTimer?.Dispose();
            Application.RemoveMessageFilter(_smoothScroller);
            _smoothScroller.Dispose();
            _showWindowRegisteredWait?.Unregister(null);
            _showWindowEvent?.Dispose();
            _activityToast?.Dispose();
            AppTheme.Changed -= ApplyTheme;
            Microsoft.Win32.SystemEvents.UserPreferenceChanged -= OnSystemThemeChanged;
        }

        base.Dispose(disposing);
    }

    private void EnsureFileAssociationRegistered()
    {
        // 매번 재등록한다(레지스트리 쓰기만 하는 가벼운 작업) — 표시 이름/아이콘/지원 확장자가 바뀐 새 빌드로
        // 교체됐을 때도 사용자가 별도로 "등록"을 누르지 않아도 최신 정보로 자동 갱신되게 하기 위함.
        try
        {
            _registrar.RegisterAll(_config, _exePath);
        }
        catch (Exception ex)
        {
            _logger.Warn($"파일 연결 자동 등록 실패: {ex.Message}");
        }
    }

    private void LoadAutoStartState()
    {
        try
        {
            var enabled = _startupRegistrar.IsAutoStartEnabled();
            var minimized = _startupRegistrar.IsStartMinimized();
            // 읽어 온 값을 스위치에 넣는 것은 사용자가 바꾼 것이 아니므로 다시 저장하지 않는다.
            _loadingAutoStartState = true;
            _autoStartCheckBox.Checked = enabled;
            _startMinimizedCheckBox.Checked = enabled && minimized;
            _startMinimizedCheckBox.Enabled = enabled;
        }
        catch (Exception ex)
        {
            _logger.Warn($"자동 시작 상태 확인 실패: {ex.Message}");
        }
        finally
        {
            _loadingAutoStartState = false;
        }
    }

    private bool _loadingAutoStartState;

    private void OnAutoStartCheckedChanged()
    {
        if (_loadingAutoStartState)
        {
            return;
        }

        _startMinimizedCheckBox.Enabled = _autoStartCheckBox.Checked;

        if (!_autoStartCheckBox.Checked && _startMinimizedCheckBox.Checked)
        {
            _startMinimizedCheckBox.Checked = false; // 이 변경 자체가 CheckedChanged를 통해 저장까지 처리함
            return;
        }

        SaveAutoStartSettings();
    }

    private void SaveAutoStartSettings()
    {
        if (_loadingAutoStartState)
        {
            return;
        }

        try
        {
            _startupRegistrar.SetAutoStart(_autoStartCheckBox.Checked, _startMinimizedCheckBox.Checked, _exePath);
            _logger.Info($"자동 시작 설정 변경: enabled={_autoStartCheckBox.Checked}, minimized={_startMinimizedCheckBox.Checked}");
        }
        catch (Exception ex)
        {
            _logger.Warn($"자동 시작 설정 저장 실패: {ex.Message}");
        }
    }

    private async Task OnSharedPcToggledAsync()
    {
        var enablingSharedPc = _sharedPcCheckBox.Checked;
        var hadCachedAccount = _hasCachedAccount || _hasCachedGoogleAccount;

        _config.SharedPcMode = enablingSharedPc;
        ConfigLoader.Save(_config);
        _logger.Info($"공용 PC 모드 변경: {enablingSharedPc}");

        if (enablingSharedPc && hadCachedAccount)
        {
            try
            {
                await _authService.SignOutAsync();
            }
            catch (Exception ex)
            {
                _logger.Warn($"공용 PC 전환 중 로그아웃 실패: {ex.Message}");
            }

            GraphAuthService.ClearPersistedCache();
            try
            {
                if (_hasCachedGoogleAccount)
                {
                    await _googleAuthService.SignOutAsync();
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"공용 PC 전환 중 Google 로그아웃 실패: {ex.Message}");
            }
            GoogleAuthService.ClearPersistedCache();

            _sharedPcNoticeLabel.Text = "저장돼 있던 로그인 정보를 삭제했습니다";
            await RefreshAccountStatusAsync();
            await RefreshGoogleAccountStatusAsync();
        }
        else
        {
            _sharedPcNoticeLabel.Text = string.Empty;
        }
    }

    private async Task RefreshAccountStatusAsync()
    {
        try
        {
            var username = await _authService.TryGetSignedInAccountAsync();
            _hasCachedAccount = username is not null;
            if (username is not null)
            {
                SetServiceStatus(_accountBadge, _accountStatusLabel, BadgeKind.Success, "연결됨", username);
            }
            else
            {
                SetServiceStatus(_accountBadge, _accountStatusLabel, BadgeKind.Neutral, "로그인 필요", "OneDrive로 문서를 열려면 로그인하세요");
            }
            _accountActionButton.Text = username is not null ? "로그아웃" : "로그인";
            _accountActionButton.IsPrimary = username is null;
        }
        catch (Exception ex)
        {
            _logger.Warn($"계정 상태 확인 실패: {ex.Message}");
            SetServiceStatus(_accountBadge, _accountStatusLabel, BadgeKind.Warning, "확인 실패", "잠시 후 다시 확인합니다");
        }
    }

    private async Task OnAccountActionAsync()
    {
        _accountActionButton.Enabled = false;
        var loggingOut = _hasCachedAccount;
        SetServiceStatus(_accountBadge, _accountStatusLabel, BadgeKind.Neutral, loggingOut ? "로그아웃 중" : "로그인 중", _accountStatusLabel.Text);

        try
        {
            if (loggingOut)
            {
                await _authService.SignOutAsync();
            }
            else
            {
                await _authService.AcquireTokenAsync();
            }

            await RefreshAccountStatusAsync();
        }
        catch (Exception ex)
        {
            _logger.Error(loggingOut ? "로그아웃 실패" : "로그인 실패", ex);
            SetServiceStatus(_accountBadge, _accountStatusLabel, BadgeKind.Danger,
                loggingOut ? "로그아웃 실패" : "로그인 실패",
                loggingOut ? "다시 시도해 주세요" : "네트워크를 확인하고 다시 시도해 주세요");
        }
        finally
        {
            _accountActionButton.Enabled = true;
        }
    }

    private async Task RefreshGoogleAccountStatusAsync()
    {
        try
        {
            _hasCachedGoogleAccount = await _googleAuthService.HasCachedAccountAsync();
            if (!_googleAuthService.IsConfigured)
            {
                SetServiceStatus(_googleBadge, _googleAccountStatusLabel, BadgeKind.Warning, "설정 필요", "config.json에 OAuth 클라이언트가 필요합니다");
                _googleAccountActionButton.Text = "설정 안내";
                _googleAccountActionButton.IsPrimary = false;
                return;
            }

            // 저장된 리프레시 토큰이 만료/취소된 상태. 토큰 파일은 남아 있어도 쓸 수 없으므로
            // "로그인됨"이 아니라 다시 로그인이 필요하다고 분명히 알린다.
            if (_googleAuthService.ReauthenticationRequired)
            {
                SetServiceStatus(_googleBadge, _googleAccountStatusLabel, BadgeKind.Danger, "다시 로그인 필요", "로그인이 만료되었습니다");
                _googleAccountActionButton.Text = "다시 로그인";
                _googleAccountActionButton.IsPrimary = true;
                return;
            }
            if (_hasCachedGoogleAccount)
            {
                SetServiceStatus(_googleBadge, _googleAccountStatusLabel, BadgeKind.Success, "연결됨", "테스트 OAuth");
            }
            else
            {
                SetServiceStatus(_googleBadge, _googleAccountStatusLabel, BadgeKind.Neutral, "로그인 필요", "Google Drive로 문서를 열려면 로그인하세요");
            }
            _googleAccountActionButton.Text = _hasCachedGoogleAccount ? "로그아웃" : "로그인";
            _googleAccountActionButton.IsPrimary = !_hasCachedGoogleAccount;
        }
        catch (Exception ex)
        {
            _logger.Warn($"Google 계정 상태 확인 실패: {ex.Message}");
            SetServiceStatus(_googleBadge, _googleAccountStatusLabel, BadgeKind.Warning, "확인 실패", "잠시 후 다시 확인합니다");
        }
    }

    private async Task OnGoogleAccountActionAsync()
    {
        if (!_googleAuthService.IsConfigured)
        {
            AppMessageDialog.Show(this,
                "Google Cloud에서 데스크톱 OAuth 클라이언트를 만든 뒤 config.json의 googleAuth.clientId와 clientSecret을 입력해 주세요.",
                AppBrand.Name, AppMessageKind.Warning);
            return;
        }

        _googleAccountActionButton.Enabled = false;
        var reauthenticating = _googleAuthService.ReauthenticationRequired;
        var loggingOut = !reauthenticating && _hasCachedGoogleAccount;
        SetServiceStatus(_googleBadge, _googleAccountStatusLabel, BadgeKind.Neutral, loggingOut ? "로그아웃 중" : "로그인 중", _googleAccountStatusLabel.Text);

        // 이 버튼은 사용자가 직접 누른 것이므로 이 흐름에서만 동의 창을 띄우도록 허용한다.
        // (백그라운드 동기화는 계속 비대화형으로 두어 갑자기 브라우저가 뜨지 않게 한다.)
        _googleAuthService.InteractiveAuthAllowed = true;
        try
        {
            if (loggingOut)
            {
                await _googleAuthService.SignOutAsync();
            }
            else if (reauthenticating)
            {
                // 만료된 토큰은 버리고 동의 창을 강제로 띄운다.
                await _googleAuthService.ForceInteractiveSignInAsync();
            }
            else
            {
                await _googleAuthService.AcquireCredentialAsync();
            }
            await RefreshGoogleAccountStatusAsync();
        }
        catch (Exception ex)
        {
            _logger.Error(loggingOut ? "Google 로그아웃 실패" : "Google 로그인 실패", ex);
            SetServiceStatus(_googleBadge, _googleAccountStatusLabel, BadgeKind.Danger,
                loggingOut ? "로그아웃 실패" : "로그인 실패", "다시 시도해 주세요");
        }
        finally
        {
            _googleAuthService.InteractiveAuthAllowed = false;
            _googleAccountActionButton.Enabled = true;
        }
    }

    private void RefreshFileAssocStatus()
    {
        var registered = _registrar.IsRegistered();
        _fileAssocBadge.Set(registered ? BadgeKind.Success : BadgeKind.Warning, registered ? "연결됨" : "연결 안 됨");
        _fileAssocActionButton.Text = registered ? "해제" : "등록";
    }

    private void OnFileAssocActionClicked()
    {
        _fileAssocActionButton.Enabled = false;
        var registering = !_registrar.IsRegistered();
        _fileAssocBadge.Set(BadgeKind.Neutral, registering ? "등록 중" : "해제 중");

        try
        {
            if (registering)
            {
                _registrar.RegisterAll(_config, _exePath);
            }
            else
            {
                _registrar.UnregisterAll(_config);
            }
        }
        catch (Exception ex)
        {
            _logger.Error("파일 연결 등록/해제 실패", ex);
        }
        finally
        {
            RefreshFileAssocStatus();
            _fileAssocActionButton.Enabled = true;
        }
    }

    private void OnOpenDefaultAppsUiClicked()
    {
        try
        {
            _registrar.OpenAdvancedAssociationUI();
        }
        catch (Exception ex)
        {
            _logger.Warn($"기본 프로그램 설정 창 열기 실패: {ex.Message}");
        }
    }

    private void RefreshSyncStatus()
    {
        var entries = _manifest.AllEntries;
        if (entries.Count == 0)
        {
            _syncStatusLabel.Text = "아직 Blue Page로 연 문서가 없습니다";
            return;
        }

        var lastSync = entries.Values.Max(e => e.LastKnownRemoteModifiedUtc > e.LastKnownLocalWriteUtc
            ? e.LastKnownRemoteModifiedUtc
            : e.LastKnownLocalWriteUtc);
        _syncStatusLabel.Text = $"문서 {entries.Count}개, 마지막 동기화 {FormatRelative(lastSync)}";
    }

    private async Task OnSyncActionAsync()
    {
        if (_syncInProgress)
        {
            // 백그라운드 동기화가 도는 중이면 끝난 직후에 검토 창을 연다(예전에는 조용히 무시했다).
            if (_backgroundSyncRunning)
            {
                _openSyncReviewAfterBackgroundSync = true;
                _syncActionButton.Enabled = false;
                _syncActionButton.Text = "동기화 중...";
            }
            return;
        }

        // 더블클릭으로 열었던 다른 프로세스가 방금 기록한 최신 상태를 놓치지 않도록 항상 다시 읽어온다.
        _manifest.Reload();

        // 로컬에서 사라진 파일도 함께 보여 줘서 검토 창에서 목록에서 뺄 수 있게 한다.
        var paths = _manifest.AllEntries.Keys.ToList();
        if (paths.Count == 0)
        {
            AppMessageDialog.Show(this, "동기화할 파일이 없습니다.", AppBrand.Name);
            return;
        }

        // 검토 창이 열려 있는 동안은 백그라운드 동기화가 같은 파일을 건드리지 않게 막는다.
        _syncInProgress = true;
        _syncActionButton.Enabled = false;

        try
        {
            // 창을 먼저 띄우고, 파일별 상태 조회는 창 안에서 한다.
            using var dialog = new SyncSelectionDialog(
                paths,
                (path, ct) => _orchestrator.DetectSyncStatusAsync(path, ct),
                (path, ex) => _logger.Warn($"동기화 상태 확인 실패: {path} — {ex.Message}"));
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                var failures = new List<string>();
                var conflictCopies = new List<string>();
                foreach (var (path, action) in dialog.SelectedActions)
                {
                    try
                    {
                        var result = await _orchestrator.ApplySyncActionAsync(path, action, CancellationToken.None);
                        if (!string.IsNullOrEmpty(result.ConflictCopyPath))
                        {
                            conflictCopies.Add(result.ConflictCopyPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error($"동기화 적용 실패: {path}", ex);
                        var reason = GraphErrorHelper.IsResourceLocked(ex)
                            ? "온라인에서 편집 중이라 반영하지 못했습니다"
                            : "반영하지 못했습니다(로그 확인)";
                        failures.Add($"• {Path.GetFileName(path)} — {reason}");
                    }
                }

                if (failures.Count > 0)
                {
                    AppMessageDialog.Show(
                        this,
                        "일부 파일을 반영하지 못했습니다:\n\n" + string.Join("\n", failures),
                        AppBrand.Name, AppMessageKind.Warning);
                }

                if (conflictCopies.Count == 1)
                {
                    LaunchOrchestrator.ShowConflictCopySaved(conflictCopies[0], this);
                }
                else if (conflictCopies.Count > 1)
                {
                    AppMessageDialog.ShowWithAction(
                        this,
                        $"충돌 사본 {conflictCopies.Count}개를 백업 폴더에 저장했습니다. 두 파일을 확인 후 직접 병합해 주세요.\n\n" +
                        string.Join("\n", conflictCopies.Select(Path.GetFileName)),
                        "폴더 열기",
                        () => ExplorerLauncher.OpenFolder(LocalBackupService.BackupRootDirectory));
                }
            }
        }
        catch (Exception ex)
        {
            _logger.Error("동기화 목록 확인 실패", ex);
        }
        finally
        {
            _syncInProgress = false;
            RefreshSyncStatus();
            _syncActionButton.Enabled = true;
            _syncActionButton.Text = SyncReviewButtonText;
        }
    }

    private async Task OnOpenOneDriveClickedAsync()
    {
        try
        {
            var webUrl = await _uploadService.GetAppFolderWebUrlAsync(CancellationToken.None);
            Process.Start(new ProcessStartInfo(webUrl) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            _logger.Warn($"OneDrive 폴더 열기 실패: {ex.Message}");
            AppMessageDialog.Show(
                this,
                "OneDrive 폴더를 여는 데 실패했습니다. 먼저 로그인이 되어 있는지 확인해 주세요.",
                AppBrand.Name,
                AppMessageKind.Warning);
        }
    }

    private async Task OnOpenGoogleDriveClickedAsync()
    {
        try
        {
            var webUrl = await _googleDriveService.GetBluePageFolderWebUrlAsync(CancellationToken.None);
            Process.Start(new ProcessStartInfo(webUrl) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            _logger.Warn($"Google Drive 폴더 열기 실패: {ex.Message}");
            AppMessageDialog.Show(this,
                "Google Drive의 BluePage 폴더를 여는 데 실패했습니다. Google OAuth 설정과 로그인을 확인해 주세요.",
                AppBrand.Name, AppMessageKind.Warning);
        }
    }

    private void OpenLogFolder()
    {
        var logDir = Path.GetDirectoryName(_logger.CurrentLogFilePath)!;
        Directory.CreateDirectory(logDir);
        Process.Start(new ProcessStartInfo("explorer.exe", $"\"{logDir}\"") { UseShellExecute = true });
    }

    private static void OpenBackupFolder() =>ExplorerLauncher.OpenFolder(LocalBackupService.BackupRootDirectory);

    private static string FormatRelative(DateTimeOffset time)
    {
        var span = DateTimeOffset.UtcNow - time;
        if (span.TotalMinutes < 1) return "방금 전";
        if (span.TotalMinutes < 60) return $"{(int)span.TotalMinutes}분 전";
        if (span.TotalHours < 24) return $"{(int)span.TotalHours}시간 전";
        return $"{(int)span.TotalDays}일 전";
    }
}
