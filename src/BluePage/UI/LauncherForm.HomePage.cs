using Microsoft365OfficeWebLauncher.UI.Controls;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>홈 화면: 서비스 계정 카드 두 개, 공용 PC 모드, 동기화 카드. 파일 연결은 설정 화면에 있다.</summary>
public sealed partial class LauncherForm
{
    private StatusBadge _accountBadge = null!;
    private Label _accountStatusLabel = null!;
    private ModernButton _accountActionButton = null!;
    private StatusBadge _googleBadge = null!;
    private Label _googleAccountStatusLabel = null!;
    private ModernButton _googleAccountActionButton = null!;
    private Label _syncStatusLabel = null!;
    private ModernButton _syncActionButton = null!;
    private CheckBox _sharedPcCheckBox = null!;
    private Label _sharedPcNoticeLabel = null!;

    private Control BuildHomePage()
    {
        var page = new PageStack("홈", "연결된 서비스와 동기화 상태를 한눈에 봅니다");

        _accountBadge = new StatusBadge();
        _accountStatusLabel = Detail("확인하는 중...");
        _accountActionButton = CreateButton("로그인");
        _accountActionButton.Click += async (_, _) => await OnAccountActionAsync();
        var openOneDrive = CreateButton("OneDrive");
        openOneDrive.Click += async (_, _) => await OnOpenOneDriveClickedAsync();

        _googleBadge = new StatusBadge();
        _googleAccountStatusLabel = Detail("확인하는 중...");
        _googleAccountActionButton = CreateButton("로그인");
        _googleAccountActionButton.Click += async (_, _) => await OnGoogleAccountActionAsync();
        var openGoogleDrive = CreateButton("Drive");
        openGoogleDrive.Click += async (_, _) => await OnOpenGoogleDriveClickedAsync();

        var services = new TableLayoutPanel { ColumnCount = 2, RowCount = 1, Height = UiDraw.S(150), Margin = new Padding(0, 0, 0, UiDraw.S(10)) };
        services.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        services.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        services.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        var microsoftCard = BuildServiceCard("M", BadgeKind.Accent, "Microsoft 365", _accountStatusLabel, _accountBadge, openOneDrive, _accountActionButton);
        microsoftCard.Margin = new Padding(0, 0, UiDraw.S(5), 0);
        var googleCard = BuildServiceCard("G", BadgeKind.Success, "Google Workspace", _googleAccountStatusLabel, _googleBadge, openGoogleDrive, _googleAccountActionButton);
        googleCard.Margin = new Padding(UiDraw.S(5), 0, 0, 0);
        services.Controls.Add(microsoftCard, 0, 0);
        services.Controls.Add(googleCard, 1, 0);
        page.Add(services);

        // 공용 PC 모드는 두 계정의 로그인 정보 저장 여부를 함께 정하므로 계정 카드 바로 아래에 둔다.
        _sharedPcCheckBox = new ToggleSwitch { Checked = _config.SharedPcMode };
        _sharedPcCheckBox.CheckedChanged += async (_, _) => await OnSharedPcToggledAsync();
        _sharedPcNoticeLabel = Detail();
        page.Add(BuildRowsCard(new Control[]
        {
            BuildSettingRow("공용 PC 모드",
                "로그인 정보를 이 PC에 저장하지 않습니다. 켜면 저장돼 있던 Microsoft와 Google 로그인 정보도 지웁니다.",
                _sharedPcCheckBox, _sharedPcNoticeLabel)
        }, SettingRowHeight));

        _syncStatusLabel = Detail("확인하는 중...");
        _syncActionButton = CreateButton(SyncReviewButtonText, primary: true);
        _syncActionButton.Click += async (_, _) => await OnSyncActionAsync();
        page.Add(BuildWideCard(new IconTile(Glyphs.Sync, BadgeKind.Warm, useIconFont: true), "동기화", _syncStatusLabel, _syncActionButton));

        return page.Root;
    }

    /// <summary>서비스 카드: 위에 아이콘과 이름, 계정. 아래에 상태 배지와 버튼.</summary>
    private static CardPanel BuildServiceCard(string letter, BadgeKind tileKind, string name, Label detail, StatusBadge badge, params Control[] buttons)
    {
        var card = new CardPanel { Dock = DockStyle.Fill, Padding = UiDraw.S(18, 16, 18, 16) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Margin = new Padding(0) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, UiDraw.S(46)));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, UiDraw.S(36)));
        card.Controls.Add(layout);

        var header = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = new Padding(0) };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, UiDraw.S(54)));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        header.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        var tile = new IconTile(letter, tileKind, useIconFont: false) { Anchor = AnchorStyles.Left };
        header.Controls.Add(tile, 0, 0);
        header.Controls.Add(BuildTextStack(name, detail, CardTitleFont), 1, 0);
        layout.Controls.Add(header, 0, 0);
        layout.Controls.Add(new Panel { Dock = DockStyle.Fill, Margin = new Padding(0), Height = 3 }, 0, 1);

        var footer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = new Padding(0) };
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        footer.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        badge.Anchor = AnchorStyles.Left;
        footer.Controls.Add(badge, 0, 0);
        footer.Controls.Add(BuildButtonRow(buttons), 1, 0);
        layout.Controls.Add(footer, 0, 2);
        return card;
    }

    /// <summary>한 줄짜리 넓은 카드: 아이콘, 제목과 설명, 오른쪽 버튼들.</summary>
    private static CardPanel BuildWideCard(IconTile tile, string title, Label detail, params Control[] buttons)
    {
        var card = new CardPanel { Height = UiDraw.S(86), Padding = UiDraw.S(18, 12, 18, 12), Margin = new Padding(0, 0, 0, UiDraw.S(10)) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, Margin = new Padding(0) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, UiDraw.S(54)));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        card.Controls.Add(layout);

        tile.Anchor = AnchorStyles.Left;
        layout.Controls.Add(tile, 0, 0);

        layout.Controls.Add(BuildTextStack(title, detail, CardTitleFont), 1, 0);

        var buttonRow = BuildButtonRow(buttons);
        buttonRow.Anchor = AnchorStyles.Right;
        buttonRow.Dock = DockStyle.None;
        layout.Controls.Add(buttonRow, 2, 0);
        return card;
    }

    private static FlowLayoutPanel BuildButtonRow(Control[] buttons)
    {
        var row = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, WrapContents = false, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0) };
        // 오른쪽에서부터 쌓이므로 마지막 버튼이 가장 오른쪽에 온다.
        for (var i = buttons.Length - 1; i >= 0; i--)
        {
            row.Controls.Add(buttons[i]);
        }
        return row;
    }
}
