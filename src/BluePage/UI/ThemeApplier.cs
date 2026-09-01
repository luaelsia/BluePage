namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>
/// 창 하나의 컨트롤 트리를 순회하며 현재 테마 색상을 입힌다. 강조색을 직접 관리하는 라벨(동기화 상태 등)은
/// Tag에 "theme-skip"을 넣어 이 자동 채색에서 제외할 수 있다.
/// </summary>
public static class ThemeApplier
{
    public const string SkipTag = "theme-skip";
    public const string SecondaryTag = "theme-secondary";
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

            case ModernCardPanel card:
                card.BackColor = theme.CardBackground;
                card.ForeColor = theme.TextPrimary;
                card.Invalidate();
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

            case NumericUpDown numeric:
                numeric.BackColor = theme.ButtonBackground;
                numeric.ForeColor = theme.TextPrimary;
                break;

            case DataGridView grid:
                ApplyToGrid(grid, theme);
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
                panel.BackColor = theme.Background;
                panel.ForeColor = theme.TextPrimary;
                break;
        }
    }

    private static void ApplyToGrid(DataGridView grid, ThemePalette theme)
    {
        grid.BackgroundColor = theme.Background;
        grid.GridColor = theme.Border;
        grid.EnableHeadersVisualStyles = false;
        grid.BorderStyle = BorderStyle.None;

        grid.ColumnHeadersDefaultCellStyle.BackColor = theme.CardBackground;
        grid.ColumnHeadersDefaultCellStyle.ForeColor = theme.TextPrimary;
        grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = theme.CardBackground;
        grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = theme.TextPrimary;

        grid.DefaultCellStyle.BackColor = theme.CardBackground;
        grid.DefaultCellStyle.ForeColor = theme.TextPrimary;
        grid.DefaultCellStyle.SelectionBackColor = theme.Accent;
        grid.DefaultCellStyle.SelectionForeColor = theme.CardBackground;

        grid.RowsDefaultCellStyle.BackColor = theme.CardBackground;
        grid.RowsDefaultCellStyle.ForeColor = theme.TextPrimary;
        grid.AlternatingRowsDefaultCellStyle.BackColor = theme.Background;
        grid.AlternatingRowsDefaultCellStyle.ForeColor = theme.TextPrimary;
    }
}
