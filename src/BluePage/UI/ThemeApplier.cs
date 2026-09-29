using Microsoft365OfficeWebLauncher.UI.Controls;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>
/// 창 하나의 컨트롤 트리를 순회하며 현재 테마 색상을 입힌다. 강조색을 직접 관리하는 라벨(동기화 상태 등)은
/// Tag에 "theme-skip"을 넣어 이 자동 채색에서 제외할 수 있다.
/// </summary>
public static class ThemeApplier
{
    public const string SkipTag = "theme-skip";
    public const string SecondaryTag = "theme-secondary";
    public const string SidebarTag = "theme-sidebar";
    public const int DialogButtonHeight = 34;

    public static void Apply(Form form, ThemePalette theme)
    {
        form.BackColor = theme.Background;
        form.ForeColor = theme.TextPrimary;
        ApplyToChildren(form.Controls, theme);

        // Handle 프로퍼티를 읽으면(아직 생성 전이어도) 핸들이 즉시 생성되므로, 여기서 바로 타이틀바까지 맞출 수 있다.
        AppTheme.ApplyTitleBarTheme(form.Handle, AppTheme.IsDark);
    }

    private static void ApplyToChildren(Control.ControlCollection controls, ThemePalette theme)
    {
        foreach (Control control in controls)
        {
            if (control.Tag as string != SkipTag)
            {
                ApplyToControl(control, theme);
            }

            if (control.HasChildren)
            {
                ApplyToChildren(control.Controls, theme);
            }
        }
    }

    private static void ApplyToControl(Control control, ThemePalette theme)
    {
        switch (control)
        {
            case TabPage tabPage:
                tabPage.BackColor = theme.Background;
                tabPage.ForeColor = theme.TextPrimary;
                break;

            case TabControl tabControl:
                tabControl.BackColor = theme.Background;
                tabControl.ForeColor = theme.TextPrimary;
                tabControl.Invalidate();
                break;

            case CardPanel card:
                card.BackColor = theme.CardBackground;
                card.ForeColor = theme.TextPrimary;
                card.Invalidate();
                break;

            case Panel sidebar when sidebar.Tag as string == SidebarTag:
                sidebar.BackColor = theme.Sidebar;
                sidebar.ForeColor = theme.TextPrimary;
                sidebar.Invalidate();
                break;

            // 직접 그리는 컨트롤은 그릴 때 AppTheme.Current를 읽으므로 다시 그리기만 하면 된다.
            case ToggleSwitch or StatusBadge or IconTile or SegmentedControl or SelectBox or NavItem or FoldCorner or WhaleFooter:
                control.Invalidate();
                break;

            case ModernButton modernButton:
                modernButton.BackColor = theme.ButtonBackground;
                modernButton.ForeColor = theme.TextPrimary;
                modernButton.Invalidate();
                break;

            case Button button:
                button.FlatStyle = FlatStyle.Flat;
                button.AutoSizeMode = AutoSizeMode.GrowOnly;
                button.MinimumSize = new Size(button.MinimumSize.Width, DialogButtonHeight);
                button.Padding = new Padding(10, 0, 10, 0);
                button.Margin = new Padding(button.Margin.Left, 0, button.Margin.Right, 0);
                button.BackColor = theme.ButtonBackground;
                button.ForeColor = theme.TextPrimary;
                button.FlatAppearance.BorderColor = theme.ButtonBorder;
                button.FlatAppearance.MouseOverBackColor = theme.ButtonHover;
                break;

            case ComboBox comboBox:
                comboBox.FlatStyle = FlatStyle.Flat;
                comboBox.BackColor = theme.ButtonBackground;
                comboBox.ForeColor = theme.TextPrimary;
                break;

            case TextBox textBox when textBox.Parent is not UpDownBase: // NumericUpDown 내부 입력칸은 제외
                textBox.BorderStyle = BorderStyle.FixedSingle;
                textBox.BackColor = theme.ButtonBackground;
                textBox.ForeColor = theme.TextPrimary;
                if (textBox.Multiline)
                {
                    SetWindowTheme(textBox.Handle, AppTheme.IsDark ? "DarkMode_Explorer" : "Explorer", null);
                }
                break;

            case NumericUpDown numeric:
                numeric.BackColor = theme.ButtonBackground;
                numeric.ForeColor = theme.TextPrimary;
                break;

            case DataGridView grid:
                ApplyToGrid(grid, theme);
                break;

            case ScrollBar scrollBar:
                SetWindowTheme(scrollBar.Handle, AppTheme.IsDark ? "DarkMode_Explorer" : "Explorer", null);
                break;

            case LinkLabel link:
                link.ForeColor = theme.Link;
                link.LinkColor = theme.Link;
                link.ActiveLinkColor = theme.Accent;
                link.VisitedLinkColor = theme.Link;
                break;

            case CheckBox or RadioButton:
                control.ForeColor = theme.TextPrimary;
                break;

            case Label label:
                label.ForeColor = label.Tag as string == SecondaryTag
                    ? theme.TextSecondary
                    : theme.TextPrimary;
                break;

            case Panel panel when panel.Height <= 2:
                // BuildDivider()로 만든 1px 구분선
                panel.BackColor = theme.Border;
                break;

            case Panel panel:
                panel.BackColor = SurfaceOf(panel, theme);
                panel.ForeColor = theme.TextPrimary;
                if (panel.AutoScroll)
                {
                    // 스크롤바를 Windows 기본 다크 테마로 맞춘다(지원하지 않는 환경에서는 아무 일도 없다).
                    SetWindowTheme(panel.Handle, AppTheme.IsDark ? "DarkMode_Explorer" : "Explorer", null);
                }
                break;
        }
    }

    [System.Runtime.InteropServices.DllImport("uxtheme.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    private static extern int SetWindowTheme(IntPtr hwnd, string? subAppName, string? subIdList);

    /// <summary>배치용 패널은 자신을 감싼 카드나 사이드바의 색을 따른다. 아무것도 없으면 창 배경색.</summary>
    private static Color SurfaceOf(Control control, ThemePalette theme)
    {
        for (var parent = control.Parent; parent is not null; parent = parent.Parent)
        {
            if (parent is CardPanel) return theme.CardBackground;
            if (parent.Tag as string == SidebarTag) return theme.Sidebar;
        }
        return theme.Background;
    }

    private static void ApplyToGrid(DataGridView grid, ThemePalette theme)
    {
        // 세로 구분선 없이 가로줄만 두고, 선택은 연한 강조색으로 표시한다.
        var surface = SurfaceOf(grid, theme);
        grid.BackgroundColor = surface;
        grid.GridColor = theme.Border;
        grid.EnableHeadersVisualStyles = false;
        grid.BorderStyle = BorderStyle.None;
        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;

        var cellPadding = UiDraw.S(10, 0, 10, 0);
        grid.ColumnHeadersDefaultCellStyle.BackColor = surface;
        grid.ColumnHeadersDefaultCellStyle.ForeColor = theme.TextSecondary;
        grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = surface;
        grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = theme.TextSecondary;
        grid.ColumnHeadersDefaultCellStyle.Font = new Font(grid.Font, FontStyle.Bold);
        grid.ColumnHeadersDefaultCellStyle.Padding = cellPadding;

        grid.DefaultCellStyle.BackColor = surface;
        grid.DefaultCellStyle.ForeColor = theme.TextPrimary;
        grid.DefaultCellStyle.SelectionBackColor = theme.AccentSoft;
        grid.DefaultCellStyle.SelectionForeColor = theme.AccentSoftText;
        grid.DefaultCellStyle.Padding = cellPadding;

        grid.RowsDefaultCellStyle.BackColor = surface;
        grid.RowsDefaultCellStyle.ForeColor = theme.TextPrimary;
        grid.AlternatingRowsDefaultCellStyle.BackColor = surface;
        grid.AlternatingRowsDefaultCellStyle.ForeColor = theme.TextPrimary;
    }
}
