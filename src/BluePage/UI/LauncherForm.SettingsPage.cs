using Microsoft365OfficeWebLauncher.Config;
using Microsoft365OfficeWebLauncher.UI.Controls;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>설정 화면: 동기화 / 화면 / 시작 / 파일 연결 / 관리 / 업데이트 묶음. 공용 PC 모드는 계정과 함께 홈에 있다.</summary>
public sealed partial class LauncherForm
{
    private static readonly int[] SyncIntervalPresets = { 10, 30, 60, 300 };

    private CheckBox _autoStartCheckBox = null!;
    private StatusBadge _fileAssocBadge = null!;
    private ModernButton _fileAssocActionButton = null!;
    private CheckBox _startMinimizedCheckBox = null!;
    private CheckBox _showToastCheckBox = null!;
    private SegmentedControl _themeSelector = null!;
    private SegmentedControl _syncIntervalSelector = null!;
    private TextBox _syncIntervalCustomInput = null!;
    private Label _syncIntervalDetail = null!;
    private ModernButton _checkUpdateButton = null!;
    private Label _updateStatusLabel = null!;
    private CheckBox _checkUpdatesCheckBox = null!;

    private Control BuildSettingsPage()
    {
        var page = new PageStack("설정", null);
        var rowHeight = SettingRowHeight;

        // 동기화
        _syncIntervalDetail = Detail();
        _showToastCheckBox = new ToggleSwitch { Checked = _config.ShowSyncToast };
        _showToastCheckBox.CheckedChanged += (_, _) => OnShowToastCheckedChanged();
        page.AddSection("동기화");
        page.Add(BuildRowsCard(new Control[]
        {
            BuildSettingRow("자동 동기화 주기", $"온라인에서 바뀐 내용을 이 간격마다 확인해 로컬 파일에 반영합니다. 직접 입력은 {MinSyncIntervalSeconds}~{MaxSyncIntervalSeconds}초, 기본값은 10초입니다.", BuildSyncIntervalControl(), _syncIntervalDetail),
            BuildSettingRow("동기화 알림", "파일을 올리거나 받을 때 화면 오른쪽 아래에 알림을 띄웁니다.", _showToastCheckBox)
        }, rowHeight));

        // 화면
        _themeSelector = new SegmentedControl("시스템", "라이트", "다크");
        _themeSelector.SelectedIndex = AppTheme.Preference switch
        {
            ThemePreference.Light => 1,
            ThemePreference.Dark => 2,
            _ => 0
        };
        _themeSelector.SelectedIndexChanged += (_, _) => OnThemeChanged();
        page.AddSection("화면");
        page.Add(BuildRowsCard(new Control[]
        {
            BuildSettingRow("테마", "시스템을 고르면 Windows의 라이트/다크 설정을 따라갑니다.", _themeSelector)
        }, rowHeight));

        // 시작
        _autoStartCheckBox = new ToggleSwitch();
        _autoStartCheckBox.CheckedChanged += (_, _) => OnAutoStartCheckedChanged();
        _startMinimizedCheckBox = new ToggleSwitch();
        _startMinimizedCheckBox.CheckedChanged += (_, _) => SaveAutoStartSettings();
        page.AddSection("시작");
        page.Add(BuildRowsCard(new Control[]
        {
            BuildSettingRow("Windows 시작 시 자동 실행", "Windows에 로그인하면 Blue Page를 자동으로 시작합니다.", _autoStartCheckBox),
            BuildSettingRow("트레이로 최소화해서 시작", "자동 실행할 때 창을 띄우지 않고 트레이 아이콘으로만 시작합니다. 자동 실행을 켰을 때만 쓸 수 있습니다.", _startMinimizedCheckBox)
        }, rowHeight));

        // 파일 연결
        _fileAssocBadge = new StatusBadge();
        _fileAssocActionButton = CreateButton("해제");
        _fileAssocActionButton.Click += (_, _) => OnFileAssocActionClicked();
        var defaultApps = CreateButton("기본 앱 설정");
        defaultApps.Click += (_, _) => OnOpenDefaultAppsUiClicked();
        var fileAssocButtons = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, WrapContents = false, AutoSize = true, Margin = new Padding(0) };
        fileAssocButtons.Controls.Add(defaultApps);
        fileAssocButtons.Controls.Add(_fileAssocActionButton);
        var formatCount = _config.DocumentTypes.Sum(type => type.Extensions.Count);
        page.AddSection("파일 연결");
        page.Add(BuildRowsCard(new Control[]
        {
            BuildSettingRow("파일 연결",
                $"탐색기의 연결 프로그램 목록에 Blue Page를 등록합니다(Office 문서 {formatCount}개 형식). [기본 앱 설정]에서 항상 Blue Page로 열도록 지정할 수 있습니다.",
                fileAssocButtons, _fileAssocBadge)
        }, rowHeight));

        // 관리
        var logButton = CreateButton("열기");
        logButton.Click += (_, _) => OpenLogFolder();
        var backupButton = CreateButton("열기");
        backupButton.Click += (_, _) => OpenBackupFolder();
        page.AddSection("관리");
        page.Add(BuildRowsCard(new Control[]
        {
            BuildSettingRow("로그 폴더", "동작 기록이 날짜별로 저장된 폴더입니다. 문제가 생겼을 때 원인을 찾는 데 씁니다.", logButton),
            BuildSettingRow("백업 폴더", "로컬 파일을 온라인 내용으로 덮어쓰기 전의 파일과 충돌 사본을 이 PC에 보관하는 폴더입니다.", backupButton)
        }, rowHeight));

        // 업데이트
        _updateStatusLabel = Detail();
        _checkUpdateButton = CreateButton("업데이트 확인");
        _checkUpdateButton.Click += async (_, _) => await CheckForUpdatesAsync(manual: true);
        _checkUpdatesCheckBox = new ToggleSwitch { Checked = _config.CheckForUpdates };
        _checkUpdatesCheckBox.CheckedChanged += (_, _) =>
        {
            _config.CheckForUpdates = _checkUpdatesCheckBox.Checked;
            ConfigLoader.Save(_config);
            _logger.Info($"업데이트 자동 확인 설정 변경: {_checkUpdatesCheckBox.Checked}");
        };
        page.AddSection("업데이트");
        page.Add(BuildRowsCard(new Control[]
        {
            BuildSettingRow($"현재 버전 {AppBrand.Version}", "GitHub Releases에서 새 버전이 있는지 확인합니다. 새 버전이 있으면 받아서 설치할 수 있습니다.", _checkUpdateButton, _updateStatusLabel),
            BuildSettingRow("새 버전 자동 확인", "시작하고 30초 뒤와 그 뒤 하루에 한 번 새 버전을 확인하고, 있으면 트레이로 알립니다.", _checkUpdatesCheckBox)
        }, rowHeight));

        return page.Root;
    }

    /// <summary>주기 프리셋(10초 / 30초 / 1분 / 5분)과 "직접" 입력. 직접 입력은 타이핑하는 대로 바로 적용한다.</summary>
    private Control BuildSyncIntervalControl()
    {
        var panel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, WrapContents = false, AutoSize = true, Margin = new Padding(0) };
        var current = Math.Clamp(_config.BackgroundSyncIntervalSeconds, MinSyncIntervalSeconds, MaxSyncIntervalSeconds);

        // Text는 이벤트를 연결하기 전에 넣는다(생성 중에는 타이머가 아직 없다).
        _syncIntervalCustomInput = new TextBox
        {
            Text = current.ToString(),
            Width = UiDraw.S(56),
            TextAlign = HorizontalAlignment.Right,
            MaxLength = 3,
            Margin = new Padding(0, UiDraw.S(5), UiDraw.S(10), 0),
            Visible = false
        };
        _syncIntervalCustomInput.KeyPress += (_, e) =>
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        };
        _syncIntervalCustomInput.TextChanged += (_, _) =>
        {
            if (int.TryParse(_syncIntervalCustomInput.Text, out var seconds) &&
                seconds is >= MinSyncIntervalSeconds and <= MaxSyncIntervalSeconds)
            {
                ApplySyncInterval(seconds);
                UpdateSyncIntervalDetail(customInvalid: false);
            }
            else
            {
                UpdateSyncIntervalDetail(customInvalid: true);
            }
        };

        _syncIntervalSelector = new SegmentedControl("10초", "30초", "1분", "5분", "직접");
        var presetIndex = Array.IndexOf(SyncIntervalPresets, current);
        _syncIntervalSelector.SelectedIndex = presetIndex >= 0 ? presetIndex : SyncIntervalPresets.Length;
        _syncIntervalCustomInput.Visible = presetIndex < 0;
        _syncIntervalSelector.SelectedIndexChanged += (_, _) =>
        {
            var index = _syncIntervalSelector.SelectedIndex;
            var custom = index >= SyncIntervalPresets.Length;
            _syncIntervalCustomInput.Visible = custom;
            if (custom)
            {
                _syncIntervalCustomInput.Text = _config.BackgroundSyncIntervalSeconds.ToString();
                _syncIntervalCustomInput.Focus();
                _syncIntervalCustomInput.SelectAll();
            }
            else
            {
                ApplySyncInterval(SyncIntervalPresets[index]);
            }
            UpdateSyncIntervalDetail(customInvalid: false);
        };

        panel.Controls.Add(_syncIntervalCustomInput);
        panel.Controls.Add(_syncIntervalSelector);
        UpdateSyncIntervalDetail(customInvalid: false);
        return panel;
    }

    private void UpdateSyncIntervalDetail(bool customInvalid)
    {
        if (customInvalid)
        {
            _syncIntervalDetail.Text = $"{MinSyncIntervalSeconds}~{MaxSyncIntervalSeconds}초 사이로 입력하세요";
            return;
        }
        _syncIntervalDetail.Text = string.Empty;
    }
}
