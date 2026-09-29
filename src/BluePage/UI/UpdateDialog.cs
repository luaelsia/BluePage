using System.Diagnostics;
using Microsoft365OfficeWebLauncher.Update;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>
/// 새 버전 안내 창. [지금 업데이트]를 누르면 설치 파일을 받고 확인까지 마친 뒤 DialogResult.OK로 닫힌다
/// (InstallerPath에 받은 파일 경로). 설치 실행과 앱 종료는 호출한 쪽이 한다.
/// [이 버전 건너뛰기]는 DialogResult.Ignore로 닫힌다.
/// </summary>
public sealed class UpdateDialog : Form
{
    private readonly UpdateInfo _update;
    private readonly Button _updateButton;
    private readonly Button _releasePageButton;
    private readonly Button _skipButton;
    private readonly Button _laterButton;
    private readonly ProgressBar _progressBar;
    private readonly Label _progressLabel;
    private CancellationTokenSource? _downloadCancellation;

    public string? InstallerPath { get; private set; }

    public UpdateDialog(UpdateInfo update)
    {
        _update = update;

        Text = $"{AppBrand.Name} 업데이트";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = true;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(560, 420);
        Font = new Font("Segoe UI Variable Text", 9.5F);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            Padding = new Padding(18)
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(layout);

        layout.Controls.Add(new Label
        {
            Text = $"새 버전 {update.DisplayVersion}을(를) 설치할 수 있습니다.",
            AutoSize = true,
            Font = new Font("Segoe UI Variable Display", 12F, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 4)
        });

        layout.Controls.Add(new Label
        {
            Text = update.InstallerUrl is null
                ? $"현재 버전 {UpdateChecker.CurrentVersion.ToString(3)}. 이 릴리스에는 설치 파일이 없어 릴리스 페이지에서 받아야 합니다."
                : $"현재 버전 {UpdateChecker.CurrentVersion.ToString(3)}. 업데이트하면 Blue Page가 잠시 종료되었다가 설치 후 다시 실행됩니다.",
            AutoSize = true,
            MaximumSize = new Size(520, 0),
            Tag = ThemeApplier.SecondaryTag,
            Margin = new Padding(0, 0, 0, 10)
        });

        layout.Controls.Add(new TextBox
        {
            Text = string.IsNullOrWhiteSpace(update.ReleaseNotes)
                ? "(릴리스 노트가 없습니다.)"
                : update.ReleaseNotes.Replace("\r\n", "\n").Replace("\n", Environment.NewLine),
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 0, 0, 10)
        });

        var progressRow = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 10)
        };
        _progressBar = new ProgressBar { Width = 380, Height = 18, Maximum = 1000, Visible = false, Margin = new Padding(0, 2, 10, 0) };
        _progressLabel = new Label { AutoSize = true, Tag = ThemeApplier.SecondaryTag, Margin = new Padding(0, 2, 0, 0) };
        progressRow.Controls.Add(_progressBar);
        progressRow.Controls.Add(_progressLabel);
        layout.Controls.Add(progressRow);

        var buttons = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
            AutoSize = true,
            Dock = DockStyle.Fill,
            Margin = new Padding(0)
        };
        _laterButton = new ModernButton { Text = "나중에", AutoSize = true };
        _laterButton.Click += (_, _) => OnLaterClicked();
        _skipButton = new ModernButton { Text = "이 버전 건너뛰기", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
        _skipButton.Click += (_, _) =>
        {
            DialogResult = DialogResult.Ignore;
            Close();
        };
        _releasePageButton = new ModernButton { Text = "릴리스 페이지", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
        _releasePageButton.Click += (_, _) =>
            Process.Start(new ProcessStartInfo(update.ReleasePageUrl) { UseShellExecute = true });
        _updateButton = new ModernButton { Text = "지금 업데이트", AutoSize = true, Margin = new Padding(0, 0, 8, 0), Enabled = update.InstallerUrl is not null, IsPrimary = true };
        _updateButton.Click += async (_, _) => await OnUpdateClickedAsync();
        buttons.Controls.Add(_laterButton);
        buttons.Controls.Add(_skipButton);
        buttons.Controls.Add(_releasePageButton);
        buttons.Controls.Add(_updateButton);
        layout.Controls.Add(buttons);

        AcceptButton = _updateButton.Enabled ? _updateButton : _releasePageButton;
        FormClosing += (_, e) =>
        {
            // 받는 중에 창을 닫으면 다운로드를 취소한다.
            _downloadCancellation?.Cancel();
        };

        ThemeApplier.Apply(this, AppTheme.Current);
    }

    private void OnLaterClicked()
    {
        if (_downloadCancellation is not null)
        {
            _downloadCancellation.Cancel();
            return;
        }
        DialogResult = DialogResult.Cancel;
        Close();
    }

    private async Task OnUpdateClickedAsync()
    {
        _updateButton.Enabled = false;
        _skipButton.Enabled = false;
        _laterButton.Text = "취소";
        _progressBar.Visible = true;
        _progressBar.Value = 0;
        _progressLabel.Text = "받는 중...";

        _downloadCancellation = new CancellationTokenSource();
        var progress = new Progress<double>(ratio =>
        {
            _progressBar.Value = (int)Math.Clamp(ratio * 1000, 0, 1000);
            _progressLabel.Text = $"{ratio * 100:0}%";
        });

        try
        {
            InstallerPath = await UpdateChecker.DownloadInstallerAsync(_update, progress, _downloadCancellation.Token);
            _progressLabel.Text = "확인 완료";
            DialogResult = DialogResult.OK;
            Close();
        }
        catch (OperationCanceledException)
        {
            ResetAfterDownload("취소했습니다.");
        }
        catch (Exception ex)
        {
            ResetAfterDownload("받지 못했습니다.");
            if (!IsDisposed)
            {
                AppMessageDialog.Show(this, $"업데이트 파일을 받지 못했습니다.\n\n{ex.Message}", AppBrand.Name, AppMessageKind.Warning);
            }
        }
        finally
        {
            _downloadCancellation?.Dispose();
            _downloadCancellation = null;
        }
    }

    private void ResetAfterDownload(string message)
    {
        if (IsDisposed)
        {
            return;
        }
        _progressBar.Visible = false;
        _progressLabel.Text = message;
        _updateButton.Enabled = true;
        _skipButton.Enabled = true;
        _laterButton.Text = "나중에";
    }
}
