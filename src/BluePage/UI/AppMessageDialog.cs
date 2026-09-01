namespace Microsoft365OfficeWebLauncher.UI;

public enum AppMessageKind
{
    Information,
    Warning,
    Error
}

/// <summary>앱의 모든 단일 확인 알림에 동일한 테마와 버튼 규격을 적용한다.</summary>
public sealed class AppMessageDialog : Form
{
    private AppMessageDialog(string message, string title, AppMessageKind kind)
    {
        Text = title;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterScreen;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        MinimumSize = new Size(340, 0);
        MaximumSize = new Size(560, 0);
        Font = new Font("Segoe UI", 9F);

        var layout = new TableLayoutPanel
        {
            Padding = new Padding(18),
            ColumnCount = 2,
            RowCount = 2,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 42));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        Controls.Add(layout);

        var icon = kind switch
        {
            AppMessageKind.Error => SystemIcons.Error,
            AppMessageKind.Warning => SystemIcons.Warning,
            _ => SystemIcons.Information
        };
        layout.Controls.Add(new PictureBox
        {
            Image = icon.ToBitmap(),
            Size = new Size(32, 32),
            SizeMode = PictureBoxSizeMode.StretchImage,
            Margin = new Padding(0, 2, 10, 0)
        }, 0, 0);

        layout.Controls.Add(new Label
        {
            Text = message,
            AutoSize = true,
            MaximumSize = new Size(450, 0),
            Margin = new Padding(0, 0, 0, 18)
        }, 1, 0);

        var buttonPanel = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Dock = DockStyle.Fill,
            Margin = new Padding(0)
        };
        layout.SetColumnSpan(buttonPanel, 2);
        layout.Controls.Add(buttonPanel, 0, 1);

        var okButton = new Button { Text = "확인", AutoSize = true, DialogResult = DialogResult.OK };
        buttonPanel.Controls.Add(okButton);
        AcceptButton = okButton;
        CancelButton = okButton;

        ThemeApplier.Apply(this, AppTheme.Current);
    }

    public static DialogResult Show(string message, string? title = null, AppMessageKind kind = AppMessageKind.Information)
    {
        using var dialog = new AppMessageDialog(message, title ?? AppBrand.Name, kind);
        return dialog.ShowDialog();
    }

    public static DialogResult Show(IWin32Window owner, string message, string? title = null, AppMessageKind kind = AppMessageKind.Information)
    {
        using var dialog = new AppMessageDialog(message, title ?? AppBrand.Name, kind)
        {
            StartPosition = FormStartPosition.CenterParent
        };
        return dialog.ShowDialog(owner);
    }
}
