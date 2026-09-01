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

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>
/// 탭이나 목록 없이, 상태(계정/파일 연결/동기화) · 설정(자동 실행/공용 PC) · 바로가기 세 그룹으로 나눈 단일 창.
/// Client ID는 빌드에 이미 포함돼 있어 이 화면에서 입력받지 않는다. 파일 연결도 실행할 때마다 자동 등록된다.
/// 인수 없이 실행하거나 --settings/--minimized로 실행했을 때만 뜨며, 문서를 여는 헤드리스 흐름에는 관여하지 않는다.
/// </summary>
public sealed class LauncherForm : Form
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

    private Label _accountStatusLabel = null!;
    private Button _accountActionButton = null!;
    private Label _googleAccountStatusLabel = null!;
    private Button _googleAccountActionButton = null!;
    private Label _fileAssocStatusLabel = null!;
    private Button _fileAssocActionButton = null!;
    private Label _syncStatusLabel = null!;
    private Button _syncActionButton = null!;
    private CheckBox _autoStartCheckBox = null!;
    private CheckBox _startMinimizedCheckBox = null!;
    private CheckBox _sharedPcCheckBox = null!;
    private Label _sharedPcNoticeLabel = null!;
    private CheckBox _showToastCheckBox = null!;
    private ComboBox _themeCombo = null!;
    private ComboBox _providerCombo = null!;
    private readonly Dictionary<string, ComboBox> _extensionProviderCombos = new(StringComparer.OrdinalIgnoreCase);
    private NumericUpDown _syncIntervalInput = null!;
    private NotifyIcon _trayIcon = null!;
    private System.Windows.Forms.Timer _backgroundSyncTimer = null!;
    private ToolStripMenuItem _trayAccountItem = null!;
    private ToolStripMenuItem _trayGoogleAccountItem = null!;
    private ToolStripMenuItem _trayFileAssocItem = null!;
    private ToolStripMenuItem _trayAutoStartItem = null!;
    private ToolStripMenuItem _trayStartMinimizedItem = null!;
    private ToolStripMenuItem _traySharedPcItem = null!;
    private ToolStripMenuItem _trayShowToastItem = null!;
    private readonly List<ToolStripMenuItem> _trayIntervalItems = new();
    private EventWaitHandle _showWindowEvent = null!;
    private RegisteredWaitHandle? _showWindowRegisteredWait;
    private SyncActivityToast _activityToast = null!;
    private ModernTabControl _mainTabs = null!;

    private const int MinSyncIntervalSeconds = 1;
    private const int MaxSyncIntervalSeconds = 600;

    private bool _hasCachedAccount;
    private bool _hasCachedGoogleAccount;
    private bool _isExiting;
    private bool _syncInProgress;

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
        BuildTrayIcon();
        BuildActivityToast();
        BuildBackgroundSyncTimer();
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
        if (_launchMinimizedToTray)
        {
            Hide();
            _trayIcon.Visible = true;
        }
    }

    private void BuildUi()
    {
        Text = AppBrand.Name;
        try
        {
            Icon = (Icon?)Icon.ExtractAssociatedIcon(Application.ExecutablePath)?.Clone();
        }
        catch
        {
            Icon = SystemIcons.Application;
        }
        ShowIcon = true;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        MinimizeBox = true;
        StartPosition = FormStartPosition.CenterScreen;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Font = new Font("Segoe UI Variable Text", 9.5F);

        var root = new TableLayoutPanel
        {
            Padding = new Padding(24, 22, 24, 20),
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Width = 720
        };
        Controls.Add(root);

        var heading = new Label
        {
            Text = "Blue Page",
            AutoSize = true,
            Font = new Font("Segoe UI Variable Display", 20F, FontStyle.Bold),
            Margin = new Padding(2, 0, 0, 2)
        };
        root.Controls.Add(heading);
        root.Controls.Add(new Label
        {
            Text = "Office 문서를 Microsoft 365 또는 Google Workspace에서 열고 동기화합니다.",
            AutoSize = true,
            ForeColor = AppTheme.Current.TextSecondary,
            Tag = ThemeApplier.SecondaryTag,
            Margin = new Padding(3, 0, 0, 14)
        });

        root.Controls.Add(BuildMainTabs());
        root.Controls.Add(BuildBrandFooter());
    }

    private ModernTabControl BuildMainTabs()
    {
        _mainTabs = new ModernTabControl
        {
            Width = 672,
            Height = 430,
            Margin = new Padding(0)
        };

        _mainTabs.AddTab("일반", BuildScrollableTabContent(BuildGeneralTab()));
        _mainTabs.AddTab("문서 연결", BuildScrollableTabContent(BuildDocumentConnectionTab()));
        _mainTabs.AddTab("설정", BuildScrollableTabContent(BuildSettingsTab()));
        return _mainTabs;
    }

    private static Control BuildScrollableTabContent(Control content)
    {
        var host = new Panel { AutoScroll = true, BackColor = AppTheme.Current.Background };
        content.Dock = DockStyle.Top;
        host.Controls.Add(content);
        return host;
    }

    private Control BuildGeneralTab()
    {
        var layout = BuildTabLayout();
        layout.Controls.Add(BuildSectionTitle("연결 상태"));
        layout.Controls.Add(BuildStatusGroup());
        layout.Controls.Add(BuildSectionTitle("바로가기", 18));
        layout.Controls.Add(BuildGeneralShortcutBar());
        return layout;
    }

    private Control BuildSettingsTab()
    {
        var layout = BuildTabLayout();
        layout.Controls.Add(BuildSectionTitle("환경 설정"));
        layout.Controls.Add(BuildSettingsGroup());
        layout.Controls.Add(BuildSectionTitle("관리", 18));
        layout.Controls.Add(BuildSettingsShortcutBar());
        return layout;
    }

    private static TableLayoutPanel BuildTabLayout() => new()
    {
        ColumnCount = 1,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        Padding = new Padding(2)
    };

    private static Label BuildSectionTitle(string text, int topMargin = 0) => new()
    {
        Text = text,
        AutoSize = true,
        Font = new Font("Segoe UI Variable Text Semibold", 11F, FontStyle.Bold),
        Margin = new Padding(0, topMargin, 0, 12)
    };

    private static Control BuildSectionCard(string title, Control content)
    {
        var layout = new TableLayoutPanel
        {
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Dock = DockStyle.Fill,
            Padding = new Padding(18, 16, 18, 17)
        };
        layout.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            Font = new Font("Segoe UI Variable Text Semibold", 11F, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 11)
        });
        content.Dock = DockStyle.Fill;
        layout.Controls.Add(content);

        var card = new ModernCardPanel
        {
            Width = 522,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0, 0, 0, 12)
        };
        card.Controls.Add(layout);
        return card;
    }

    private static Button CreateModernButton(string text, int minimumWidth = 92)
    {
        var button = new ModernButton
        {
            Text = text,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            MinimumSize = new Size(minimumWidth, 34),
            Padding = new Padding(11, 3, 11, 3),
            Margin = new Padding(0, 3, 0, 3),
            UseVisualStyleBackColor = false
        };
        return button;
    }

    private TableLayoutPanel BuildStatusGroup()
    {
        var status = new TableLayoutPanel
        {
            ColumnCount = 3,
            RowCount = 4,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };
        status.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        status.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        status.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        status.Controls.Add(new Label { Text = "Microsoft", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, 0);
        _accountStatusLabel = new Label { Text = "확인 중…", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) };
        status.Controls.Add(_accountStatusLabel, 1, 0);
        _accountActionButton = CreateModernButton("지금 로그인");
        _accountActionButton.Click += async (_, _) => await OnAccountActionAsync();
        status.Controls.Add(_accountActionButton, 2, 0);

        status.Controls.Add(new Label { Text = "Google (테스트)", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, 1);
        _googleAccountStatusLabel = new Label { Text = "확인 중…", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) };
        status.Controls.Add(_googleAccountStatusLabel, 1, 1);
        _googleAccountActionButton = CreateModernButton("지금 로그인");
        _googleAccountActionButton.Click += async (_, _) => await OnGoogleAccountActionAsync();
        status.Controls.Add(_googleAccountActionButton, 2, 1);

        status.Controls.Add(new Label { Text = "파일 연결", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, 2);
        _fileAssocStatusLabel = new Label { Text = "-", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) };
        status.Controls.Add(_fileAssocStatusLabel, 1, 2);
        _fileAssocActionButton = CreateModernButton("해제");
        _fileAssocActionButton.Click += (_, _) => OnFileAssocActionClicked();
        status.Controls.Add(_fileAssocActionButton, 2, 2);

        status.Controls.Add(new Label { Text = "동기화", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, 3);
        _syncStatusLabel = new Label { Text = "-", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) };
        status.Controls.Add(_syncStatusLabel, 1, 3);
        _syncActionButton = CreateModernButton("동기화 검토…");
        _syncActionButton.Click += async (_, _) => await OnSyncActionAsync();
        status.Controls.Add(_syncActionButton, 2, 3);

        return status;
    }

    private FlowLayoutPanel BuildSettingsGroup()
    {
        var settings = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0)
        };

        _autoStartCheckBox = new CheckBox { Text = "Windows 시작 시 자동 실행", AutoSize = true, Margin = new Padding(0, 2, 0, 2) };
        _autoStartCheckBox.CheckedChanged += (_, _) => OnAutoStartCheckedChanged();
        settings.Controls.Add(_autoStartCheckBox);

        _startMinimizedCheckBox = new CheckBox { Text = "트레이로 최소화해서 시작", AutoSize = true, Margin = new Padding(24, 2, 0, 6) };
        _startMinimizedCheckBox.CheckedChanged += (_, _) => SaveAutoStartSettings();
        settings.Controls.Add(_startMinimizedCheckBox);

        _sharedPcCheckBox = new CheckBox { Text = "공용 PC (로그인 정보 저장 안 함)", AutoSize = true, Checked = _config.SharedPcMode, Margin = new Padding(0, 2, 0, 2) };
        _sharedPcCheckBox.CheckedChanged += async (_, _) => await OnSharedPcToggledAsync();
        settings.Controls.Add(_sharedPcCheckBox);

        _sharedPcNoticeLabel = new Label { Text = string.Empty, AutoSize = true, ForeColor = AppTheme.Current.Success, Margin = new Padding(0, 2, 0, 2), Visible = false, Tag = ThemeApplier.SkipTag };
        settings.Controls.Add(_sharedPcNoticeLabel);

        _showToastCheckBox = new CheckBox { Text = "동기화 알림 표시(화면 우하단)", AutoSize = true, Checked = _config.ShowSyncToast, Margin = new Padding(0, 2, 0, 2) };
        _showToastCheckBox.CheckedChanged += (_, _) => OnShowToastCheckedChanged();
        settings.Controls.Add(_showToastCheckBox);

        settings.Controls.Add(BuildThemeRow());
        settings.Controls.Add(BuildSyncIntervalRow());

        return settings;
    }

    private Control BuildDocumentConnectionTab()
    {
        var root = BuildTabLayout();
        root.Controls.Add(BuildSectionTitle("기본 열기 방식"));

        var defaultRow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, WrapContents = false, AutoSize = true, Margin = new Padding(0, 0, 0, 18) };
        defaultRow.Controls.Add(new Label { Text = "모든 문서:", AutoSize = true, Width = 150, Margin = new Padding(0, 7, 8, 0) });
        _providerCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 190 };
        _providerCombo.Items.AddRange(new object[] { "열 때마다 선택", "Microsoft 365", "Google Workspace (테스트)" });
        _providerCombo.SelectedIndex = _config.PreferredCloudProvider.ToLowerInvariant() switch
        {
            "microsoft" => 1,
            "google" => 2,
            _ => 0
        };
        _providerCombo.SelectedIndexChanged += (_, _) =>
        {
            _config.PreferredCloudProvider = _providerCombo.SelectedIndex switch { 1 => "Microsoft", 2 => "Google", _ => "Ask" };
            ConfigLoader.Save(_config);
        };
        defaultRow.Controls.Add(_providerCombo);
        root.Controls.Add(defaultRow);

        root.Controls.Add(BuildSectionTitle("확장자별 열기 방식"));
        root.Controls.Add(new Label
        {
            Text = "기본값과 다른 서비스로 열 확장자만 개별 지정할 수 있습니다.",
            AutoSize = true,
            Tag = ThemeApplier.SecondaryTag,
            Margin = new Padding(0, 0, 0, 12)
        });

        var grid = new TableLayoutPanel { ColumnCount = 3, AutoSize = true, CellBorderStyle = TableLayoutPanelCellBorderStyle.Single };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250));
        grid.Controls.Add(CreateGridHeader("확장자"), 0, 0);
        grid.Controls.Add(CreateGridHeader("문서 형식"), 1, 0);
        grid.Controls.Add(CreateGridHeader("열기 방식"), 2, 0);

        var definitions = _config.DocumentTypes
            .SelectMany(type => type.Extensions.Select(extension => (Extension: NormalizeExtension(extension), type.OfficeApp)))
            .OrderBy(item => item.OfficeApp)
            .ThenBy(item => item.Extension)
            .ToList();
        var rowIndex = 1;
        foreach (var definition in definitions)
        {
            grid.Controls.Add(CreateGridCell(definition.Extension), 0, rowIndex);
            grid.Controls.Add(CreateGridCell(definition.OfficeApp), 1, rowIndex);

            var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220, Margin = new Padding(8, 5, 8, 5) };
            combo.Items.AddRange(new object[] { "기본값 따름", "열 때마다 선택", "Microsoft 365", "Google Workspace (테스트)" });
            var stored = _config.DocumentProviderPreferences.TryGetValue(definition.Extension, out var value) ? value : "Default";
            combo.SelectedIndex = stored.ToLowerInvariant() switch
            {
                "ask" => 1,
                "microsoft" => 2,
                "google" => 3,
                _ => 0
            };
            var extension = definition.Extension;
            combo.SelectedIndexChanged += (_, _) =>
            {
                var preference = combo.SelectedIndex switch { 1 => "Ask", 2 => "Microsoft", 3 => "Google", _ => "Default" };
                if (preference == "Default") _config.DocumentProviderPreferences.Remove(extension);
                else _config.DocumentProviderPreferences[extension] = preference;
                ConfigLoader.Save(_config);
            };
            _extensionProviderCombos[extension] = combo;
            grid.Controls.Add(combo, 2, rowIndex++);
        }
        root.Controls.Add(grid);
        return root;
    }

    private static Label CreateGridHeader(string text) => new()
    {
        Text = text,
        AutoSize = false,
        Dock = DockStyle.Fill,
        Height = 32,
        TextAlign = ContentAlignment.MiddleLeft,
        Font = new Font("Segoe UI Variable Text", 9.5F, FontStyle.Bold),
        Padding = new Padding(8, 0, 0, 0),
        Margin = new Padding(0)
    };

    private static Label CreateGridCell(string text) => new()
    {
        Text = text,
        AutoSize = false,
        Dock = DockStyle.Fill,
        Height = 36,
        TextAlign = ContentAlignment.MiddleLeft,
        Padding = new Padding(8, 0, 0, 0),
        Margin = new Padding(0)
    };

    private static string NormalizeExtension(string extension) =>
        (extension.StartsWith('.') ? extension : "." + extension).ToLowerInvariant();

    private FlowLayoutPanel BuildThemeRow()
    {
        var row = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            AutoSize = true,
            Margin = new Padding(0, 6, 0, 2)
        };

        row.Controls.Add(new Label { Text = "테마:", AutoSize = true, Margin = new Padding(0, 4, 6, 0) });

        _themeCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
        _themeCombo.Items.AddRange(new object[] { "시스템 설정 따라가기", "라이트", "다크" });
        _themeCombo.SelectedIndex = AppTheme.Preference switch
        {
            ThemePreference.Light => 1,
            ThemePreference.Dark => 2,
            _ => 0
        };
        _themeCombo.SelectedIndexChanged += (_, _) => OnThemeChanged();
        row.Controls.Add(_themeCombo);

        return row;
    }

    private void OnThemeChanged()
    {
        var preference = _themeCombo.SelectedIndex switch
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

    /// <summary>수동 전환(위 콤보) 또는 Windows 시스템 테마 변경(AppTheme.Changed) 시 창과 트레이 메뉴를 다시 칠한다.</summary>
    private void ApplyTheme()
    {
        var theme = AppTheme.Current;
        ThemeApplier.Apply(this, theme);
        _mainTabs?.ApplyTheme(theme);

        if (_trayIcon.ContextMenuStrip is { } menu)
        {
            menu.Renderer = new ToolStripProfessionalRenderer(new ThemedMenuColorTable(theme));
            menu.ForeColor = theme.TextPrimary;
            foreach (ToolStripItem item in menu.Items)
            {
                item.ForeColor = theme.TextPrimary;
            }
        }
    }

    private void OnSystemThemeChanged(object? sender, Microsoft.Win32.UserPreferenceChangedEventArgs e) =>
        AppTheme.NotifySystemThemeChanged();

    private sealed class ThemedMenuColorTable : ProfessionalColorTable
    {
        private readonly ThemePalette _theme;
        public ThemedMenuColorTable(ThemePalette theme) => _theme = theme;

        public override Color ToolStripDropDownBackground => _theme.CardBackground;
        public override Color ImageMarginGradientBegin => _theme.CardBackground;
        public override Color ImageMarginGradientMiddle => _theme.CardBackground;
        public override Color ImageMarginGradientEnd => _theme.CardBackground;
        public override Color MenuItemSelected => _theme.ButtonHover;
        public override Color MenuItemSelectedGradientBegin => _theme.ButtonHover;
        public override Color MenuItemSelectedGradientEnd => _theme.ButtonHover;
        public override Color MenuItemBorder => _theme.Accent;
        public override Color MenuBorder => _theme.Border;
        public override Color SeparatorDark => _theme.Border;
        public override Color SeparatorLight => _theme.Border;
    }

    private FlowLayoutPanel BuildSyncIntervalRow()
    {
        var row = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            AutoSize = true,
            Margin = new Padding(0, 6, 0, 2)
        };

        row.Controls.Add(new Label { Text = "자동 동기화 주기(초):", AutoSize = true, Margin = new Padding(0, 4, 6, 0) });

        _syncIntervalInput = new NumericUpDown
        {
            Minimum = MinSyncIntervalSeconds,
            Maximum = MaxSyncIntervalSeconds,
            Value = Math.Clamp(_config.BackgroundSyncIntervalSeconds, MinSyncIntervalSeconds, MaxSyncIntervalSeconds),
            Width = 70
        };
        _syncIntervalInput.ValueChanged += (_, _) => OnSyncIntervalChanged();
        row.Controls.Add(_syncIntervalInput);

        row.Controls.Add(new Label
        {
            Text = $"({MinSyncIntervalSeconds}~{MaxSyncIntervalSeconds}초, 기본 180초)",
            AutoSize = true,
            ForeColor = AppTheme.Current.TextSecondary,
            Tag = ThemeApplier.SecondaryTag,
            Margin = new Padding(6, 4, 0, 0)
        });

        return row;
    }

    private FlowLayoutPanel BuildGeneralShortcutBar()
    {
        var bar = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };

        var setDefaultButton = CreateModernButton("기본 프로그램으로 설정…", 150);
        setDefaultButton.Margin = new Padding(0, 0, 8, 0);
        setDefaultButton.Click += (_, _) => OnOpenDefaultAppsUiClicked();
        bar.Controls.Add(setDefaultButton);

        var openOneDriveButton = CreateModernButton("OneDrive에서 보기…", 130);
        openOneDriveButton.Margin = new Padding(0, 0, 8, 0);
        openOneDriveButton.Click += async (_, _) => await OnOpenOneDriveClickedAsync();
        bar.Controls.Add(openOneDriveButton);

        var openGoogleDriveButton = CreateModernButton("Google Drive에서 보기…", 145);
        openGoogleDriveButton.Margin = new Padding(0, 0, 8, 0);
        openGoogleDriveButton.Click += async (_, _) => await OnOpenGoogleDriveClickedAsync();
        bar.Controls.Add(openGoogleDriveButton);

        return bar;
    }

    private FlowLayoutPanel BuildSettingsShortcutBar()
    {
        var bar = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, WrapContents = false, AutoSize = true };
        var logButton = CreateModernButton("로그 폴더 열기", 120);
        logButton.Margin = new Padding(0, 0, 8, 0);
        logButton.Click += (_, _) => OpenLogFolder();
        bar.Controls.Add(logButton);
        var backupButton = CreateModernButton("백업 폴더 열기", 120);
        backupButton.Margin = new Padding(0);
        backupButton.Click += (_, _) => OpenBackupFolder();
        bar.Controls.Add(backupButton);
        return bar;
    }

    private Control BuildBrandFooter()
    {
        var footer = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0, 14, 0, 0),
            Anchor = AnchorStyles.Left
        };

        footer.Controls.Add(new Label
        {
            Text = AppBrand.Publisher,
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold),
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 3, 10, 0)
        });

        footer.Controls.Add(new Label
        {
            Text = $"v{AppBrand.Version}",
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Tag = ThemeApplier.SecondaryTag,
            Margin = new Padding(0, 3, 10, 0)
        });

        var contact = new LinkLabel
        {
            Text = AppBrand.ContactEmail,
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 3, 0, 0),
            AccessibleDescription = "MiniWhaleLabs 이메일 문의"
        };
        contact.LinkClicked += (_, _) =>
            Process.Start(new ProcessStartInfo($"mailto:{AppBrand.ContactEmail}") { UseShellExecute = true });
        footer.Controls.Add(contact);

        return footer;
    }

    private void BuildTrayIcon()
    {
        var menu = new ContextMenuStrip();

        menu.Items.Add("열기", null, (_, _) => ShowFromTray());
        menu.Items.Add(new ToolStripSeparator());

        _trayAccountItem = new ToolStripMenuItem("지금 로그인");
        _trayAccountItem.Click += async (_, _) => await OnAccountActionAsync();
        menu.Items.Add(_trayAccountItem);

        _trayGoogleAccountItem = new ToolStripMenuItem("Google 로그인 (테스트)");
        _trayGoogleAccountItem.Click += async (_, _) => await OnGoogleAccountActionAsync();
        menu.Items.Add(_trayGoogleAccountItem);

        _trayFileAssocItem = new ToolStripMenuItem("파일 연결 등록");
        _trayFileAssocItem.Click += (_, _) => OnFileAssocActionClicked();
        menu.Items.Add(_trayFileAssocItem);

        menu.Items.Add("기본 프로그램으로 설정…", null, (_, _) => OnOpenDefaultAppsUiClicked());
        menu.Items.Add(new ToolStripSeparator());

        menu.Items.Add("동기화 검토…", null, async (_, _) => await OnSyncActionAsync());
        menu.Items.Add("OneDrive에서 보기…", null, async (_, _) => await OnOpenOneDriveClickedAsync());
        menu.Items.Add("Google Drive에서 보기…", null, async (_, _) => await OnOpenGoogleDriveClickedAsync());

        var intervalMenu = new ToolStripMenuItem("자동 동기화 주기");
        foreach (var seconds in new[] { 30, 60, 180, 300, 600 })
        {
            var item = new ToolStripMenuItem(FormatIntervalLabel(seconds)) { Tag = seconds };
            item.Click += (_, _) => _syncIntervalInput.Value = seconds;
            _trayIntervalItems.Add(item);
            intervalMenu.DropDownItems.Add(item);
        }
        menu.Items.Add(intervalMenu);
        menu.Items.Add(new ToolStripSeparator());

        _trayAutoStartItem = new ToolStripMenuItem("Windows 시작 시 자동 실행") { CheckOnClick = false };
        _trayAutoStartItem.Click += (_, _) => _autoStartCheckBox.Checked = !_autoStartCheckBox.Checked;
        menu.Items.Add(_trayAutoStartItem);

        _trayStartMinimizedItem = new ToolStripMenuItem("   트레이로 최소화해서 시작") { CheckOnClick = false };
        _trayStartMinimizedItem.Click += (_, _) => _startMinimizedCheckBox.Checked = !_startMinimizedCheckBox.Checked;
        menu.Items.Add(_trayStartMinimizedItem);

        _traySharedPcItem = new ToolStripMenuItem("공용 PC 모드") { CheckOnClick = false };
        _traySharedPcItem.Click += (_, _) => _sharedPcCheckBox.Checked = !_sharedPcCheckBox.Checked;
        menu.Items.Add(_traySharedPcItem);

        _trayShowToastItem = new ToolStripMenuItem("동기화 알림 표시") { CheckOnClick = false };
        _trayShowToastItem.Click += (_, _) => _showToastCheckBox.Checked = !_showToastCheckBox.Checked;
        menu.Items.Add(_trayShowToastItem);
        menu.Items.Add(new ToolStripSeparator());

        menu.Items.Add("로그 폴더 열기", null, (_, _) => OpenLogFolder());
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("종료", null, (_, _) => ExitApplication());

        menu.Opening += (_, _) => RefreshTrayMenuState();

        _trayIcon = new NotifyIcon
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? SystemIcons.Application,
            Text = AppBrand.Name,
            Visible = true, // 창이 보이는 상태에서도 항상 트레이에 표시(프로그램이 켜져 있는 동안 계속)
            ContextMenuStrip = menu
        };
        _trayIcon.DoubleClick += (_, _) => ShowFromTray();
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

    /// <summary>트레이 메뉴를 열 때마다 현재 상태를 반영한다 — 원본 상태는 항상 창의 컨트롤들이 갖고 있으므로 거기서 읽어온다.</summary>
    private void RefreshTrayMenuState()
    {
        _trayAccountItem.Text = _hasCachedAccount ? "로그아웃" : "지금 로그인";
        _trayGoogleAccountItem.Text = _hasCachedGoogleAccount ? "Google 로그아웃 (테스트)" : "Google 로그인 (테스트)";

        var registered = _registrar.IsRegistered();
        _trayFileAssocItem.Text = registered ? "파일 연결 해제" : "파일 연결 등록";

        _trayAutoStartItem.Checked = _autoStartCheckBox.Checked;
        _trayStartMinimizedItem.Checked = _startMinimizedCheckBox.Checked;
        _trayStartMinimizedItem.Enabled = _autoStartCheckBox.Checked;
        _traySharedPcItem.Checked = _sharedPcCheckBox.Checked;
        _trayShowToastItem.Checked = _showToastCheckBox.Checked;

        foreach (var item in _trayIntervalItems)
        {
            item.Checked = (int)item.Tag! == (int)_syncIntervalInput.Value;
        }
    }

    private static string FormatIntervalLabel(int seconds)
    {
        var label = seconds < 60 ? $"{seconds}초" : $"{seconds / 60}분";
        return seconds == 180 ? $"{label} (기본)" : label;
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

    private void OnSyncIntervalChanged()
    {
        var seconds = (int)_syncIntervalInput.Value;
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
        _logger.Info("백그라운드 자동 동기화 시작");

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
            if (!IsDisposed)
            {
                RefreshSyncStatus();
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
            _autoStartCheckBox.Checked = enabled;
            _startMinimizedCheckBox.Checked = enabled && minimized;
            _startMinimizedCheckBox.Enabled = enabled;
        }
        catch (Exception ex)
        {
            _logger.Warn($"자동 시작 상태 확인 실패: {ex.Message}");
        }
    }

    private void OnAutoStartCheckedChanged()
    {
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

            _sharedPcNoticeLabel.Text = "저장돼 있던 로그인 정보를 삭제했습니다.";
            _sharedPcNoticeLabel.ForeColor = AppTheme.Current.Success;
            _sharedPcNoticeLabel.Visible = true;
            await RefreshAccountStatusAsync();
            await RefreshGoogleAccountStatusAsync();
        }
        else
        {
            _sharedPcNoticeLabel.Visible = false;
        }
    }

    private async Task RefreshAccountStatusAsync()
    {
        try
        {
            var username = await _authService.TryGetSignedInAccountAsync();
            _hasCachedAccount = username is not null;
            _accountStatusLabel.Text = username is not null ? $"● 로그인됨 — {username}" : "○ 로그인 안 됨";
            _accountActionButton.Text = username is not null ? "로그아웃" : "지금 로그인";
        }
        catch (Exception ex)
        {
            _logger.Warn($"계정 상태 확인 실패: {ex.Message}");
            _accountStatusLabel.Text = "확인 실패";
        }
    }

    private async Task OnAccountActionAsync()
    {
        _accountActionButton.Enabled = false;
        var loggingOut = _hasCachedAccount;
        _accountStatusLabel.Text = loggingOut ? "로그아웃하는 중…" : "로그인하는 중…";

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
            _accountStatusLabel.Text = loggingOut
                ? "로그아웃에 실패했습니다."
                : "로그인에 실패했습니다. 네트워크를 확인하고 다시 시도해 주세요.";
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
                _googleAccountStatusLabel.Text = "OAuth 설정 필요 · 테스트";
                _googleAccountActionButton.Text = "설정 확인";
                return;
            }
            _googleAccountStatusLabel.Text = _hasCachedGoogleAccount ? "● 로그인됨 · 테스트 OAuth" : "○ 로그인 안 됨 · 테스트 OAuth";
            _googleAccountActionButton.Text = _hasCachedGoogleAccount ? "로그아웃" : "지금 로그인";
        }
        catch (Exception ex)
        {
            _logger.Warn($"Google 계정 상태 확인 실패: {ex.Message}");
            _googleAccountStatusLabel.Text = "확인 실패";
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
        var loggingOut = _hasCachedGoogleAccount;
        _googleAccountStatusLabel.Text = loggingOut ? "로그아웃하는 중…" : "로그인하는 중…";
        try
        {
            if (loggingOut)
            {
                await _googleAuthService.SignOutAsync();
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
            _googleAccountStatusLabel.Text = "Google 계정 처리 실패";
        }
        finally
        {
            _googleAccountActionButton.Enabled = true;
        }
    }

    private void RefreshFileAssocStatus()
    {
        var registered = _registrar.IsRegistered();
        _fileAssocStatusLabel.Text = registered ? "✅ 연결됨" : "⚠ 연결 안 됨";
        _fileAssocActionButton.Text = registered ? "해제" : "등록";
    }

    private void OnFileAssocActionClicked()
    {
        _fileAssocActionButton.Enabled = false;
        var registering = !_registrar.IsRegistered();
        _fileAssocStatusLabel.Text = registering ? "등록하는 중…" : "해제하는 중…";

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
            _syncStatusLabel.Text = "아직 연 문서가 없습니다.";
            return;
        }

        var lastSync = entries.Values.Max(e => e.LastKnownRemoteModifiedUtc > e.LastKnownLocalWriteUtc
            ? e.LastKnownRemoteModifiedUtc
            : e.LastKnownLocalWriteUtc);
        _syncStatusLabel.Text = $"{entries.Count}개 파일 · {FormatRelative(lastSync)}";
    }

    private async Task OnSyncActionAsync()
    {
        if (_syncInProgress)
        {
            return;
        }

        // 더블클릭으로 열었던 다른 프로세스가 방금 기록한 최신 상태를 놓치지 않도록 항상 다시 읽어온다.
        _manifest.Reload();

        var paths = _manifest.AllEntries.Keys.Where(File.Exists).ToList();
        if (paths.Count == 0)
        {
            AppMessageDialog.Show(this, "동기화할 파일이 없습니다.", AppBrand.Name);
            return;
        }

        _syncInProgress = true;
        _syncActionButton.Enabled = false;
        _syncStatusLabel.Text = "확인하는 중…";

        try
        {
            var detections = new List<SyncDetection>();
            foreach (var path in paths)
            {
                try
                {
                    detections.Add(await _orchestrator.DetectSyncStatusAsync(path, CancellationToken.None));
                }
                catch (Exception ex)
                {
                    _logger.Warn($"동기화 상태 확인 실패: {path} — {ex.Message}");
                }
            }

            using var dialog = new SyncSelectionDialog(detections);
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                var failures = new List<string>();
                foreach (var (path, action) in dialog.SelectedActions)
                {
                    try
                    {
                        await _orchestrator.ApplySyncActionAsync(path, action, CancellationToken.None);
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

    private static void OpenBackupFolder()
    {
        Directory.CreateDirectory(LocalBackupService.BackupRootDirectory);
        Process.Start(new ProcessStartInfo("explorer.exe", $"\"{LocalBackupService.BackupRootDirectory}\"") { UseShellExecute = true });
    }

    private static string FormatRelative(DateTimeOffset time)
    {
        var span = DateTimeOffset.UtcNow - time;
        if (span.TotalMinutes < 1) return "방금 전";
        if (span.TotalMinutes < 60) return $"{(int)span.TotalMinutes}분 전";
        if (span.TotalHours < 24) return $"{(int)span.TotalHours}시간 전";
        return $"{(int)span.TotalDays}일 전";
    }
}
