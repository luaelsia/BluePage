namespace Microsoft365OfficeWebLauncher.UI.Controls;

/// <summary>트레이 메뉴와 SelectBox 목록이 함께 쓰는 테마 색 메뉴.</summary>
internal static class ThemedMenu
{
    public static void Apply(ContextMenuStrip menu)
    {
        var theme = AppTheme.Current;
        menu.Renderer = new ToolStripProfessionalRenderer(new ColorTable(theme)) { RoundedEdges = true };
        menu.BackColor = theme.CardBackground;
        menu.ForeColor = theme.TextPrimary;
        foreach (ToolStripItem item in menu.Items)
        {
            item.ForeColor = theme.TextPrimary;
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
