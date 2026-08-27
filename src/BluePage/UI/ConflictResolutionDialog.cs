using Microsoft365OfficeWebLauncher.OneDrive;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>
/// 로컬 파일과 온라인(Office Web) 사본이 모두 변경됐을 때 4가지 처리 방식 중 하나를 사용자가 고르는 모달 창.
/// 기본 선택은 "사본 생성"(가장 안전) — 아무것도 선택하지 않고 창을 닫아도 이 기본값이 적용된다.
/// </summary>
public sealed class ConflictResolutionDialog : Form
{
    private readonly RadioButton _createCopyOption;
    private readonly RadioButton _keepNewerOption;
    private readonly RadioButton _keepLocalOption;
    private readonly RadioButton _keepRemoteOption;

    public ConflictResolutionChoice SelectedChoice { get; private set; } = ConflictResolutionChoice.CreateCopy;

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
        _keepNewerOption = new RadioButton { Text = "더 최신 파일로 덮어쓰기 — 나중에 수정된 쪽을 사용", AutoSize = true, Margin = new Padding(0, 4, 0, 4) };
        _keepLocalOption = new RadioButton { Text = "오프라인(로컬) 파일로 덮어쓰기 — 온라인 사본을 로컬 내용으로 교체", AutoSize = true, Margin = new Padding(0, 4, 0, 4) };
        _keepRemoteOption = new RadioButton { Text = "온라인 파일로 덮어쓰기 — 로컬 파일을 온라인 내용으로 교체", AutoSize = true, Margin = new Padding(0, 4, 0, 10) };

        layout.Controls.Add(_createCopyOption);
        layout.Controls.Add(_keepNewerOption);
        layout.Controls.Add(_keepLocalOption);
        layout.Controls.Add(_keepRemoteOption);

        var okButton = new Button { Text = "확인", AutoSize = true, Anchor = AnchorStyles.Right };
        okButton.Click += (_, _) =>
        {
            SelectedChoice = GetSelectedChoice();
            DialogResult = DialogResult.OK;
            Close();
        };
        AcceptButton = okButton;
        layout.Controls.Add(okButton);

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
