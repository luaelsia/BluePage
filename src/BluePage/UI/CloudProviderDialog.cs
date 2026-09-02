using Microsoft365OfficeWebLauncher.Cloud;

namespace Microsoft365OfficeWebLauncher.UI;

public sealed class CloudProviderDialog : Form
{
    public CloudProvider? SelectedProvider { get; private set; }
    public bool RememberAsDefault { get; private set; }

    public CloudProviderDialog(string filePath, IReadOnlySet<CloudProvider> supportedProviders)
    {
        Text = "웹 Office 선택";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        ShowInTaskbar = true;
        Font = new Font("Segoe UI Variable Text", 9.5F);

        var root = new TableLayoutPanel
        {
            ColumnCount = 1,
            AutoSize = true,
            Padding = new Padding(22),
            MinimumSize = new Size(420, 0)
        };
        root.Controls.Add(new Label
        {
            Text = Path.GetFileName(filePath),
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 6)
        });
        root.Controls.Add(new Label
        {
            Text = "어느 웹 Office에서 여시겠습니까?",
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 16)
        });

        var buttons = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        var microsoft = new ModernButton { Text = "Microsoft 365", MinimumSize = new Size(170, 42) };
        var google = new ModernButton { Text = "Google Workspace (테스트)", MinimumSize = new Size(190, 42) };
        microsoft.Click += (_, _) => Select(CloudProvider.Microsoft);
        google.Click += (_, _) => Select(CloudProvider.Google);
        if (supportedProviders.Contains(CloudProvider.Microsoft)) buttons.Controls.Add(microsoft);
        if (supportedProviders.Contains(CloudProvider.Google)) buttons.Controls.Add(google);
        root.Controls.Add(buttons);

        var remember = new CheckBox
        {
            Text = "앞으로 이 서비스를 기본으로 사용",
            AutoSize = true,
            Margin = new Padding(0, 14, 0, 0)
        };
        remember.CheckedChanged += (_, _) => RememberAsDefault = remember.Checked;
        root.Controls.Add(remember);
        Controls.Add(root);
        ThemeApplier.Apply(this, AppTheme.Current);
    }

    private void Select(CloudProvider provider)
    {
        SelectedProvider = provider;
        DialogResult = DialogResult.OK;
        Close();
    }
}
