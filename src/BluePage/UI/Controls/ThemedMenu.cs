namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>트레이 메뉴와 SelectBox 목록이 함께 쓰는 테마 색 메뉴.</summary>
internal static class ThemedMenu
{
    /// <param name="centerText">true면 항목 글자를 메뉴 폭의 가운데에 그린다(가운데 정렬한 드롭다운의 목록용).</param>
    public static void Apply(ContextMenuStrip menu, bool centerText = false)
    {
        var theme = AppTheme.Current;
        menu.Renderer = centerText
            ? new CenteredTextRenderer(new ColorTable(theme)) { RoundedEdges = true }
            : new ToolStripProfessionalRenderer(new ColorTable(theme)) { RoundedEdges = true };
        menu.BackColor = theme.CardBackground;
        menu.ForeColor = theme.TextPrimary;
        foreach (ToolStripItem item in menu.Items)
        {
            item.ForeColor = theme.TextPrimary;
        }
    }

    /// <summary>WinForms 메뉴는 TextAlign을 무시하고 글자를 왼쪽부터 그리므로, 글자 칸을 항목 전체 폭으로 넓혀 가운데에 그린다.</summary>
    private sealed class CenteredTextRenderer : ToolStripProfessionalRenderer
    {
        public CenteredTextRenderer(ProfessionalColorTable table) : base(table)
        {
        }

        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
        {
            if (e.Item is ToolStripMenuItem)
            {
                e.TextRectangle = new Rectangle(0, e.TextRectangle.Y, e.Item.Width, e.TextRectangle.Height);
                e.TextFormat = (e.TextFormat & ~(TextFormatFlags.Left | TextFormatFlags.Right)) | TextFormatFlags.HorizontalCenter;
            }
            base.OnRenderItemText(e);
        }
    }

    private sealed class ColorTable : ProfessionalColorTable
    {
        private readonly ThemePalette _theme;
        public ColorTable(ThemePalette theme) => _theme = theme;

        public override Color ToolStripDropDownBackground => _theme.CardBackground;
        public override Color ImageMarginGradientBegin => _theme.CardBackground;
        public override Color ImageMarginGradientMiddle => _theme.CardBackground;
        public override Color ImageMarginGradientEnd => _theme.CardBackground;
        public override Color MenuItemSelected => _theme.ButtonHover;
        public override Color MenuItemSelectedGradientBegin => _theme.ButtonHover;
        public override Color MenuItemSelectedGradientEnd => _theme.ButtonHover;
        public override Color MenuItemBorder => _theme.ButtonHover;
        public override Color MenuBorder => _theme.Border;
        public override Color SeparatorDark => _theme.Border;
        public override Color SeparatorLight => _theme.Border;
        public override Color CheckBackground => _theme.AccentSoft;
        public override Color CheckSelectedBackground => _theme.AccentSoft;
        public override Color CheckPressedBackground => _theme.AccentSoft;
    }
}
