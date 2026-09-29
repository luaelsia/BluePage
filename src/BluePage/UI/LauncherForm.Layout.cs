using System.Diagnostics;
using Microsoft365OfficeWebLauncher.UI.Controls;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>창의 뼈대(사이드 메뉴, 화면 전환)와 화면들이 함께 쓰는 카드, 행, 버튼 만들기.</summary>
public sealed partial class LauncherForm
{
    private const string SyncReviewButtonText = "동기화 검토";

    private readonly List<(NavItem Nav, Control Page)> _pages = new();
    private readonly SmoothScroller _smoothScroller = new();
    private readonly List<ScrollableControl> _scrollablePages = new();
    private bool _scrollFadesCreated;

    /// <summary>
    /// 스크롤 화면 위아래의 흐림 띠를 만든다(보통 세기 30px).
    /// 레이어드 자식 창은 부모 창이 실제로 만들어진 뒤에만 만들 수 있어서, 창이 처음 뜬 뒤(Shown)에 부른다.
    /// </summary>
    private void CreateScrollFades()
    {
        if (_scrollFadesCreated)
        {
            return;
        }
        _scrollFadesCreated = true;
        foreach (var page in _scrollablePages)
        {
            _ = new ScrollFades(page, _pageHost, _smoothScroller, UiDraw.S(30));
        }
    }
    private Panel _pageHost = null!;

    private static readonly Font TitleFont = new("Segoe UI Variable Display", 17F, FontStyle.Bold);
    private static readonly Font CardTitleFont = new("Segoe UI Variable Text", 10.5F, FontStyle.Bold);
    private static readonly Font SectionFont = new("Segoe UI Variable Text", 9F, FontStyle.Bold);
    private static readonly Font SmallFont = new("Segoe UI Variable Text", 8.75F);

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
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimizeBox = true;
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI Variable Text", 9.5F);
        Size = new Size(UiDraw.S(920), UiDraw.S(660));
        MinimumSize = new Size(UiDraw.S(780), UiDraw.S(540));

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = new Padding(0), Padding = new Padding(0) };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, SidebarWidth));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        Controls.Add(root);

        _pageHost = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0) };
        root.Controls.Add(BuildSidebar(), 0, 0);
        root.Controls.Add(_pageHost, 1, 0);

        AddPage(Glyphs.Home, "홈", BuildHomePage());
        AddPage(Glyphs.Document, "문서 연결", BuildDocumentsPage());
        AddPage(Glyphs.Settings, "설정", BuildSettingsPage());
        ShowPage(0);
    }

    private TableLayoutPanel _navList = null!;

    private static int SidebarWidth => UiDraw.S(208);

    private Control BuildSidebar()
    {
        // 사이드바는 여백 없이 두고, 메뉴 영역에만 여백을 준다. 아래쪽 고래 영역이 사이드바 폭 끝까지 닿게 하기 위해서다.
        var sidebar = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0), Padding = new Padding(0), Tag = ThemeApplier.SidebarTag };
        var menuArea = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0), Padding = UiDraw.S(14, 20, 14, 0) };
        sidebar.Controls.Add(menuArea);
        sidebar.Controls.Add(BuildSidebarFooter());

        // 오른쪽 위를 아이콘처럼 접는다(아이콘 비율의 약 55%). 메뉴 위에 올라오도록 맨 앞으로 보낸다.
        var fold = new FoldCorner(0.55F * SidebarWidth / IconShapes.BaseWidth);
        sidebar.Controls.Add(fold);
        fold.BringToFront();
        sidebar.Resize += (_, _) => fold.Location = new Point(sidebar.ClientSize.Width - fold.Width, 0);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Margin = new Padding(0) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        menuArea.Controls.Add(layout);

        // 앱 아이콘과 이름
        var brand = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Margin = UiDraw.S(6, 0, 0, 18) };
        brand.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, UiDraw.S(40)));
        brand.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        var logo = new PictureBox
        {
            Size = new Size(UiDraw.S(30), UiDraw.S(30)),
            SizeMode = PictureBoxSizeMode.Zoom,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0),
            Image = Icon?.ToBitmap()
        };
        brand.Controls.Add(logo, 0, 0);
        brand.Controls.Add(new Label
        {
            Text = AppBrand.Name,
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Font = new Font("Segoe UI Variable Display", 12.5F, FontStyle.Bold),
            Margin = new Padding(0)
        }, 1, 0);
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, UiDraw.S(52)));
        layout.Controls.Add(brand);

        _navList = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, AutoSize = true, Margin = new Padding(0) };
        _navList.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(_navList);

        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        layout.Controls.Add(new Panel { Dock = DockStyle.Fill, Margin = new Padding(0) });

        return sidebar;
    }

    /// <summary>사이드바 맨 아래: 옅은 고래 곡선 위에 버전, 제작자, 연락처.</summary>
    private static Control BuildSidebarFooter()
    {
        // 고래(폭에 맞춰 늘어남) 높이와 그 아래 몸통 영역을 합친 높이. 글자는 몸통 영역 안에 들어간다.
        var whaleHeight = (int)Math.Ceiling(IconShapes.WhaleHeight * SidebarWidth / IconShapes.BaseWidth);
        var footer = new WhaleFooter { Dock = DockStyle.Bottom, Margin = new Padding(0), Padding = UiDraw.S(22, 0, 14, 14) };
        footer.Height = whaleHeight + footer.BodyBelowWhale + UiDraw.S(12);

        // 글자 뒤로 고래 곡선이 보이도록 글자 쪽 배경은 투명하게 둔다(ThemeApplier가 칠하지 않게 SkipTag).
        var text = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoSize = true,
            Dock = DockStyle.Bottom,
            BackColor = Color.Transparent,
            Margin = new Padding(0),
            Tag = ThemeApplier.SkipTag
        };
        text.Controls.Add(new Label { Text = $"{AppBrand.Name} {AppBrand.Version}", AutoSize = true, BackColor = Color.Transparent, Tag = ThemeApplier.SecondaryTag, Font = SmallFont, Margin = new Padding(0, 0, 0, UiDraw.S(2)) });
        text.Controls.Add(new Label { Text = AppBrand.Publisher, AutoSize = true, BackColor = Color.Transparent, Tag = ThemeApplier.SecondaryTag, Font = SmallFont, Margin = new Padding(0, 0, 0, UiDraw.S(2)) });
        var contact = new LinkLabel { Text = AppBrand.ContactEmail, AutoSize = true, BackColor = Color.Transparent, Font = SmallFont, Margin = new Padding(0), LinkBehavior = LinkBehavior.HoverUnderline, AccessibleDescription = "MiniWhaleLabs 이메일 문의" };
        contact.LinkClicked += (_, _) => Process.Start(new ProcessStartInfo($"mailto:{AppBrand.ContactEmail}") { UseShellExecute = true });
        text.Controls.Add(contact);
        footer.Controls.Add(text);
        return footer;
    }

    private void AddPage(string glyph, string title, Control page)
    {
        var index = _pages.Count;
        var nav = new NavItem(glyph, title) { Dock = DockStyle.Fill };
        nav.Click += (_, _) => ShowPage(index);
        _navList.RowStyles.Add(new RowStyle(SizeType.Absolute, nav.Height + nav.Margin.Vertical));
        _navList.Controls.Add(nav);

        page.Dock = DockStyle.Fill;
        page.Visible = false;
        if (page is ScrollableControl { AutoScroll: true } scrollable)
        {
            _smoothScroller.Register(scrollable);
            _scrollablePages.Add(scrollable);
        }
        _pageHost.Controls.Add(page);
        _pages.Add((nav, page));
    }

    private void ShowPage(int index)
    {
        for (var i = 0; i < _pages.Count; i++)
        {
            _pages[i].Nav.Selected = i == index;
            _pages[i].Page.Visible = i == index;
        }
    }

    // ---------- 화면 공통 부품 ----------

    /// <summary>스크롤되는 화면 하나. 제목과 설명 아래로 Add(...)한 부품이 위에서부터 쌓인다.</summary>
    private sealed class PageStack
    {
        public Panel Root { get; }
        private readonly TableLayoutPanel _stack;

        public PageStack(string title, string? subtitle)
        {
            // 끝까지 스크롤했을 때 마지막 카드가 창 바닥에 붙지 않도록 아래에 여유를 둔다.
            // 자동 스크롤 범위는 패널의 Padding을 넣지 않고 안쪽 내용의 끝까지만 계산하므로,
            // 여유는 패널이 아니라 내용 묶음(_stack)의 아래 Padding으로 준다.
            Root = new Panel { AutoScroll = true, Padding = UiDraw.S(32, 26, 32, 0), Margin = new Padding(0) };
            _stack = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, Margin = new Padding(0), Padding = new Padding(0, 0, 0, UiDraw.S(100)) };
            _stack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            Root.Controls.Add(_stack);

            Add(new Label { Text = title, AutoSize = true, Font = TitleFont, Margin = new Padding(0, 0, 0, UiDraw.S(string.IsNullOrEmpty(subtitle) ? 12 : 2)) });
            if (!string.IsNullOrEmpty(subtitle))
            {
                Add(new Label { Text = subtitle, AutoSize = true, Tag = ThemeApplier.SecondaryTag, Margin = new Padding(0, 0, 0, UiDraw.S(18)) });
            }
        }

        public void Add(Control control)
        {
            _stack.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            if (!control.AutoSize)
            {
                control.Dock = DockStyle.Fill;
            }
            _stack.Controls.Add(control);
        }

        public void AddSection(string caption) =>
            Add(new Label { Text = caption, AutoSize = true, Font = SectionFont, Tag = ThemeApplier.SecondaryTag, Margin = UiDraw.S(4, 8, 0, 6) });
    }

    private static ModernButton CreateButton(string text, bool primary = false)
    {
        return new ModernButton
        {
            Text = text,
            IsPrimary = primary,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            MinimumSize = new Size(UiDraw.S(76), UiDraw.S(34)),
            Padding = UiDraw.S(14, 0, 14, 0),
            Margin = new Padding(UiDraw.S(8), 0, 0, 0),
            UseVisualStyleBackColor = false
        };
    }

    /// <summary>제목 한 줄과 흐린 설명 한 줄. 설명이 비어 있으면 제목만 세로 가운데에 둔다.</summary>
    private static TableLayoutPanel BuildTextStack(string title, Label detail, Font? titleFont = null)
    {
        var stack = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Margin = new Padding(0) };
        stack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        stack.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
        stack.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
        var titleLabel = new Label
        {
            Text = title,
            AutoSize = true,
            Font = titleFont,
            Anchor = AnchorStyles.Left | AnchorStyles.Bottom,
            Margin = new Padding(0, 0, 0, UiDraw.S(1))
        };
        detail.AutoSize = false;
        detail.AutoEllipsis = true;
        detail.Dock = DockStyle.Fill;
        detail.Tag = ThemeApplier.SecondaryTag;
        detail.Margin = new Padding(0, UiDraw.S(1), 0, 0);
        detail.TextAlign = ContentAlignment.TopLeft;
        stack.Controls.Add(titleLabel, 0, 0);
        stack.Controls.Add(detail, 0, 1);
        return stack;
    }

    /// <summary>카드 안에 행들을 1px 구분선으로 나눠 쌓는다. 행 높이는 모두 같다.</summary>
    private static CardPanel BuildRowsCard(IReadOnlyList<Control> rows, int rowHeight)
    {
        var card = new CardPanel { Padding = UiDraw.S(20, 4, 20, 4), Margin = new Padding(0, 0, 0, UiDraw.S(10)) };
        var table = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Margin = new Padding(0) };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        for (var i = 0; i < rows.Count; i++)
        {
            if (i > 0)
            {
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 1));
                table.Controls.Add(new Panel { Height = 1, Dock = DockStyle.Fill, Margin = new Padding(0) });
            }
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, rowHeight));
            rows[i].Dock = DockStyle.Fill;
            table.Controls.Add(rows[i]);
        }
        card.Controls.Add(table);
        card.Height = rows.Count * rowHeight + (rows.Count - 1) + card.Padding.Vertical;
        return card;
    }

    private readonly ThemedToolTip _toolTip = new();

    /// <summary>설정 행의 높이. 설명을 툴팁으로 옮겨 한 줄로 줄였다.</summary>
    private static int SettingRowHeight => UiDraw.S(52);

    /// <summary>
    /// 설정 한 줄: 왼쪽에 제목과 ⓘ(설명 툴팁), 필요하면 흐린 상태 문구, 오른쪽에 스위치나 버튼 같은 컨트롤.
    /// 설명은 제목, ⓘ, 행 어디에 마우스를 올려도 보인다.
    /// </summary>
    private TableLayoutPanel BuildSettingRow(string title, string description, Control control, Control? status = null)
    {
        var row = new TableLayoutPanel { ColumnCount = 2, RowCount = 1, Margin = new Padding(0) };
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        row.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var titleRow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, WrapContents = false, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0) };
        var titleLabel = new Label { Text = title, AutoSize = true, Margin = new Padding(0, 0, UiDraw.S(4), 0) };
        var info = new Label
        {
            Text = Glyphs.Info,
            AutoSize = true,
            Font = UiDraw.IconFont(8.5F),
            Tag = ThemeApplier.SecondaryTag,
            Cursor = Cursors.Help,
            Margin = new Padding(0, UiDraw.S(3), UiDraw.S(10), 0),
            AccessibleName = $"{title} 설명",
            AccessibleDescription = description
        };
        titleRow.Controls.Add(titleLabel);
        titleRow.Controls.Add(info);
        if (status is Label statusLabel)
        {
            statusLabel.AutoSize = true;
            statusLabel.Tag = ThemeApplier.SecondaryTag;
            statusLabel.Font = SmallFont;
            statusLabel.Margin = new Padding(0, UiDraw.S(1), 0, 0);
            titleRow.Controls.Add(statusLabel);
        }
        else if (status is not null)
        {
            // 배지처럼 스스로 그리는 상태 표시는 크기와 색을 그대로 둔다.
            status.Margin = new Padding(0, UiDraw.S(1), 0, 0);
            titleRow.Controls.Add(status);
        }
        row.Controls.Add(titleRow, 0, 0);

        foreach (var target in new Control[] { row, titleRow, titleLabel, info })
        {
            _toolTip.SetToolTip(target, description);
        }

        control.Anchor = AnchorStyles.Right;
        control.Margin = new Padding(UiDraw.S(16), 0, 0, 0);
        row.Controls.Add(control, 1, 0);
        return row;
    }

    private static Label Detail(string text = "") => new() { Text = text };

    private static void SetServiceStatus(StatusBadge badge, Label detail, BadgeKind kind, string badgeText, string detailText)
    {
        badge.Set(kind, badgeText);
        detail.Text = detailText;
    }
}
