using Microsoft365OfficeWebLauncher.OneDrive;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>
/// "동기화 검토…" 클릭 시 실제로 업로드/다운로드하기 전에 보여주는 검토 창.
/// 파일마다 감지된 상태에 맞는 동작만 드롭다운으로 고를 수 있고, 기본값은 이미 합리적으로 선택돼 있다.
/// "건너뛰기"를 고른 파일은 [선택 항목 동기화]를 눌러도 전혀 처리되지 않는다.
/// </summary>
public sealed class SyncSelectionDialog : Form
{
    private sealed record ActionOption(string Display, SyncAction Action)
    {
        public override string ToString() => Display;
    }

    private readonly DataGridView _grid;

    public IReadOnlyList<(string Path, SyncAction Action)> SelectedActions { get; private set; } = Array.Empty<(string, SyncAction)>();

    public SyncSelectionDialog(IReadOnlyList<SyncDetection> detections)
    {
        Text = "동기화 검토";
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        MaximizeBox = true;
        StartPosition = FormStartPosition.CenterScreen;
        Size = new Size(680, 380);
        Font = new Font("Segoe UI", 9F);

        _grid = BuildGrid();
        Controls.Add(_grid);

        var buttonPanel = BuildButtonPanel();
        Controls.Add(buttonPanel);

        PopulateRows(detections);

        ThemeApplier.Apply(this, AppTheme.Current);
    }

    private DataGridView BuildGrid()
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.CellSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            EditMode = DataGridViewEditMode.EditOnEnter
        };

        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "File", HeaderText = "파일", ReadOnly = true, FillWeight = 32 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Status", HeaderText = "상태", ReadOnly = true, FillWeight = 18 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "LastSynced", HeaderText = "마지막 동기화", ReadOnly = true, FillWeight = 20 });
        grid.Columns.Add(new DataGridViewComboBoxColumn
        {
            Name = "Action",
            HeaderText = "실행할 동작",
            FillWeight = 30,
            DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
        });

        return grid;
    }

    private FlowLayoutPanel BuildButtonPanel()
    {
        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            FlowDirection = FlowDirection.RightToLeft,
            Height = ThemeApplier.DialogButtonHeight + 16,
            Padding = new Padding(8)
        };

        var cancelButton = new Button { Text = "취소", AutoSize = true, DialogResult = DialogResult.Cancel };
        panel.Controls.Add(cancelButton);

        var runButton = new Button { Text = "선택 항목 동기화", AutoSize = true };
        runButton.Click += (_, _) => OnRunClicked();
        panel.Controls.Add(runButton);

        AcceptButton = runButton;
        CancelButton = cancelButton;

        return panel;
    }

    private void PopulateRows(IReadOnlyList<SyncDetection> detections)
    {
        foreach (var detection in detections)
        {
            var rowIndex = _grid.Rows.Add();
            var row = _grid.Rows[rowIndex];

            row.Tag = detection.LocalFilePath;
            row.Cells["File"].Value = Path.GetFileName(detection.LocalFilePath);
            row.Cells["Status"].Value = StatusLabel(detection.DetectedState);
            row.Cells["LastSynced"].Value = FormatRelative(detection.LastSyncedUtc);

            var (options, defaultAction) = OptionsFor(detection.DetectedState);
            var comboCell = (DataGridViewComboBoxCell)row.Cells["Action"];
            comboCell.DataSource = options;
            comboCell.DisplayMember = nameof(ActionOption.Display);
            comboCell.ValueMember = nameof(ActionOption.Action);
            comboCell.Value = defaultAction;

            if (detection.DetectedState == SyncState.NoChange)
            {
                row.Cells["Action"].ReadOnly = true;
                row.DefaultCellStyle.ForeColor = AppTheme.Current.TextSecondary;
            }
        }
    }

    private void OnRunClicked()
    {
        _grid.EndEdit();

        var selected = new List<(string Path, SyncAction Action)>();
        foreach (DataGridViewRow row in _grid.Rows)
        {
            var action = (SyncAction)((DataGridViewComboBoxCell)row.Cells["Action"]).Value;
            if (action == SyncAction.Skip)
            {
                continue;
            }

            selected.Add(((string)row.Tag!, action));
        }

        SelectedActions = selected;
        DialogResult = DialogResult.OK;
        Close();
    }

    private static (List<ActionOption> Options, SyncAction Default) OptionsFor(SyncState state) => state switch
    {
        SyncState.RemoteOnlyChanged => (
            new List<ActionOption>
            {
                new("건너뛰기", SyncAction.Skip),
                new("온라인 → 로컬 반영", SyncAction.PullRemoteToLocal)
            },
            SyncAction.PullRemoteToLocal),

        SyncState.LocalOnlyChanged => (
            new List<ActionOption>
            {
                new("건너뛰기", SyncAction.Skip),
                new("로컬 → 온라인 반영", SyncAction.PushLocalToRemote)
            },
            SyncAction.PushLocalToRemote),

        SyncState.Conflict => (
            new List<ActionOption>
            {
                new("건너뛰기", SyncAction.Skip),
                new("사본 생성", SyncAction.CreateConflictCopy),
                new("온라인 → 로컬 반영", SyncAction.PullRemoteToLocal),
                new("로컬 → 온라인 반영", SyncAction.PushLocalToRemote)
            },
            SyncAction.CreateConflictCopy),

        _ => (new List<ActionOption> { new("건너뛰기", SyncAction.Skip) }, SyncAction.Skip)
    };

    private static string StatusLabel(SyncState state) => state switch
    {
        SyncState.NoChange => "변경 없음",
        SyncState.RemoteOnlyChanged => "온라인만 변경",
        SyncState.LocalOnlyChanged => "로컬만 변경",
        SyncState.Conflict => "충돌",
        _ => state.ToString()
    };

    private static string FormatRelative(DateTimeOffset time)
    {
        var span = DateTimeOffset.UtcNow - time;
        if (span.TotalMinutes < 1) return "방금 전";
        if (span.TotalMinutes < 60) return $"{(int)span.TotalMinutes}분 전";
        if (span.TotalHours < 24) return $"{(int)span.TotalHours}시간 전";
        return $"{(int)span.TotalDays}일 전";
    }
}
