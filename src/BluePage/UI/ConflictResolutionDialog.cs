using Microsoft365OfficeWebLauncher.OneDrive;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>
/// 로컬 파일과 온라인(Office Web) 사본이 모두 변경됐을 때 처리 방식을 사용자가 고르는 모달 창.
/// 창을 닫거나 "동기화 안 함"을 누르면 어느 쪽 파일도 변경하지 않는다.
/// </summary>
public sealed class ConflictResolutionDialog : Form
{
    private readonly RadioButton _createCopyOption;
    private readonly RadioButton _keepNewerOption;
    private readonly RadioButton _keepLocalOption;
    private readonly RadioButton _keepRemoteOption;

    public ConflictResolutionChoice SelectedChoice { get; private set; } = ConflictResolutionChoice.Skip;

    public ConflictResolutionDialog(ConflictInfo info)
    {
        Text = "동기화 충돌";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        TopMost = true;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Font = new Font("Segoe UI", 9F);

        var fileName = Path.GetFileName(info.LocalFilePath);
        var layout = new TableLayoutPanel
        {
            Padding = new Padding(16),
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };
        Controls.Add(layout);

        var message = new Label
        {
            Text = $"'{fileName}'의 로컬 파일과 온라인(Office Web) 사본이 모두 변경되어 자동으로 병합할 수 없습니다.\n\n" +
                   $"로컬 수정: {info.LocalModifiedUtc.ToLocalTime():yyyy-MM-dd HH:mm}\n" +
                   $"온라인 수정: {info.RemoteModifiedUtc.ToLocalTime():yyyy-MM-dd HH:mm}\n\n" +
                   "어떻게 처리할까요?",
            AutoSize = true,
            MaximumSize = new Size(400, 0),
            Margin = new Padding(0, 0, 0, 10)
        };
        layout.Controls.Add(message);

        _createCopyOption = new RadioButton { Text = "사본 생성 (안전, 권장) — 온라인 사본을 별도 파일로 저장하고 둘 다 보존", AutoSize = true, Checked = true, Margin = new Padding(0, 4, 0, 4) };
        _keepNewerOption = new RadioButton { Text = "더 최신 파일로 덮어쓰기 — 수정 시간만 비교하므로 웹을 방금 닫았다면 주의", AutoSize = true, Margin = new Padding(0, 4, 0, 4) };
        _keepLocalOption = new RadioButton { Text = "오프라인(로컬) 파일로 덮어쓰기 — 온라인 사본을 로컬 내용으로 교체", AutoSize = true, Margin = new Padding(0, 4, 0, 4) };
        _keepRemoteOption = new RadioButton { Text = "온라인 파일로 덮어쓰기 — 로컬 파일을 온라인 내용으로 교체", AutoSize = true, Margin = new Padding(0, 4, 0, 10) };

        layout.Controls.Add(_createCopyOption);
        layout.Controls.Add(_keepNewerOption);
        layout.Controls.Add(_keepLocalOption);
        layout.Controls.Add(_keepRemoteOption);

        layout.Controls.Add(new Label
        {
            Text = "온라인 파일로 로컬을 덮어쓰게 되면 현재 로컬 파일은 백업 폴더에 먼저 보관됩니다.",
            AutoSize = true,
            MaximumSize = new Size(430, 0),
            Tag = ThemeApplier.SecondaryTag,
            Margin = new Padding(0, 8, 0, 10)
        });

        var buttonPanel = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Anchor = AnchorStyles.Right,
            Margin = new Padding(0)
        };

        var okButton = new Button { Text = "선택한 방식으로 동기화", AutoSize = true };
        okButton.Click += (_, _) =>
        {
            SelectedChoice = GetSelectedChoice();
            DialogResult = DialogResult.OK;
            Close();
        };

        var skipButton = new Button { Text = "동기화 안 함", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
        skipButton.Click += (_, _) =>
        {
            SelectedChoice = ConflictResolutionChoice.Skip;
            DialogResult = DialogResult.Cancel;
            Close();
        };

        buttonPanel.Controls.Add(okButton);
        buttonPanel.Controls.Add(skipButton);
        AcceptButton = okButton;
        CancelButton = skipButton;
        layout.Controls.Add(buttonPanel);

        ThemeApplier.Apply(this, AppTheme.Current);
    }

    private ConflictResolutionChoice GetSelectedChoice()
    {
        if (_keepNewerOption.Checked) return ConflictResolutionChoice.KeepNewer;
        if (_keepLocalOption.Checked) return ConflictResolutionChoice.KeepLocal;
        if (_keepRemoteOption.Checked) return ConflictResolutionChoice.KeepRemote;
        return ConflictResolutionChoice.CreateCopy;
    }
}

/// <summary>ConflictResolutionDialog를 띄워 사용자의 선택을 받아오는 IConflictResolver 구현.</summary>
public sealed class GuiConflictResolver : IConflictResolver
{
    public ConflictResolutionChoice Resolve(ConflictInfo info)
    {
        using var dialog = new ConflictResolutionDialog(info);
        dialog.ShowDialog();
        return dialog.SelectedChoice;
    }
}
