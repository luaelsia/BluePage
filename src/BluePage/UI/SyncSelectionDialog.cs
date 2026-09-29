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

    /// <summary>상태 필터 항목. States가 null이면 전체.</summary>
    private sealed record StatusFilter(string Display, SyncState[]? States)
    {
        public override string ToString() => Display;
    }

    private static readonly StatusFilter[] StatusFilters =
    {
        new("전체", null),
        new("변경 있음", new[] { SyncState.RemoteOnlyChanged, SyncState.LocalOnlyChanged, SyncState.Conflict }),
        new("충돌", new[] { SyncState.Conflict }),
        new("온라인만 변경", new[] { SyncState.RemoteOnlyChanged }),
        new("로컬만 변경", new[] { SyncState.LocalOnlyChanged }),
        new("로컬에 없음 / 드라이브 연결 안 됨", new[] { SyncState.LocalMissing, SyncState.DriveMissing }),
        new("변경 없음", new[] { SyncState.NoChange })
    };

    private readonly DataGridView _grid;
    private ComboBox _statusFilter = null!;
    private TextBox _nameFilter = null!;
    private Label _filterCountLabel = null!;

    /// <summary>필터로 숨긴 행은 포함하지 않는다(화면에 보이는 행만 실행).</summary>
    public IReadOnlyList<(string Path, SyncAction Action)> SelectedActions { get; private set; } = Array.Empty<(string, SyncAction)>();

    public SyncSelectionDialog(IReadOnlyList<SyncDetection> detections)
    {
        Text = "동기화 검토";
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        MaximizeBox = true;
        StartPosition = FormStartPosition.CenterScreen;
        Size = new Size(960, 600);
        MinimumSize = new Size(760, 420);
        Font = new Font("Segoe UI", 9.5F);

        _grid = BuildGrid();
        Controls.Add(_grid);

        Controls.Add(BuildFilterPanel());

        var buttonPanel = BuildButtonPanel();
        Controls.Add(buttonPanel);

        PopulateRows(detections);
        ApplyFilter();

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
            EditMode = DataGridViewEditMode.EditOnEnter,
            ColumnHeadersHeight = 34,
            RowTemplate = { Height = 30 }
        };

        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "File", HeaderText = "파일", ReadOnly = true, FillWeight = 38 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Status", HeaderText = "상태", ReadOnly = true, FillWeight = 16 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "LastSynced", HeaderText = "마지막 동기화", ReadOnly = true, FillWeight = 18 });
        grid.Columns.Add(new DataGridViewComboBoxColumn
        {
            Name = "Action",
            HeaderText = "실행할 동작",
            FillWeight = 28,
            DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
        });

        return grid;
    }

    private FlowLayoutPanel BuildFilterPanel()
    {
        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Height = 44,
            Padding = new Padding(8, 8, 8, 4)
        };

        panel.Controls.Add(new Label { Text = "상태:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        _statusFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 230, Margin = new Padding(0, 2, 16, 0) };
        _statusFilter.Items.AddRange(StatusFilters);
        _statusFilter.SelectedIndex = 0;
        _statusFilter.SelectedIndexChanged += (_, _) => ApplyFilter();
        panel.Controls.Add(_statusFilter);

        panel.Controls.Add(new Label { Text = "파일 이름:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        _nameFilter = new TextBox { Width = 240, Margin = new Padding(0, 3, 16, 0), PlaceholderText = "이름 일부 입력" };
        _nameFilter.TextChanged += (_, _) => ApplyFilter();
        panel.Controls.Add(_nameFilter);

        _filterCountLabel = new Label { AutoSize = true, Tag = ThemeApplier.SecondaryTag, Margin = new Padding(0, 6, 0, 0) };
        panel.Controls.Add(_filterCountLabel);

        return panel;
    }

    private void ApplyFilter()
    {
        var states = (_statusFilter.SelectedItem as StatusFilter)?.States;
        var keyword = _nameFilter.Text.Trim();

        // 현재 셀이 있는 행은 숨길 수 없으므로 먼저 편집과 현재 셀을 해제한다.
        _grid.EndEdit();
        _grid.CurrentCell = null;

        var visibleCount = 0;
        foreach (DataGridViewRow row in _grid.Rows)
        {
            var detection = (SyncDetection)row.Tag!;
            var visible =
                (states is null || states.Contains(detection.DetectedState)) &&
                (keyword.Length == 0 ||
                 Path.GetFileName(detection.LocalFilePath).Contains(keyword, StringComparison.CurrentCultureIgnoreCase));
            row.Visible = visible;
            if (visible)
            {
                visibleCount++;
            }
        }

        _filterCountLabel.Text = visibleCount == _grid.Rows.Count
            ? $"{_grid.Rows.Count}개"
            : $"{_grid.Rows.Count}개 중 {visibleCount}개 표시 (숨긴 파일은 실행하지 않음)";
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

            row.Tag = detection;
            row.Cells["File"].Value = Path.GetFileName(detection.LocalFilePath);
            row.Cells["File"].ToolTipText = detection.LocalFilePath;
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
            if (!row.Visible)
            {
                continue;
            }

            var action = (SyncAction)((DataGridViewComboBoxCell)row.Cells["Action"]).Value;
            if (action == SyncAction.Skip)
            {
                continue;
            }

            selected.Add((((SyncDetection)row.Tag!).LocalFilePath, action));
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

        // 드라이브는 있는데 파일만 없으면 지워진 것으로 보고 빼기를 기본값으로 둔다.
        SyncState.LocalMissing => (
            new List<ActionOption>
            {
                new("건너뛰기", SyncAction.Skip),
                new("목록에서 빼기", SyncAction.RemoveFromList)
            },
            SyncAction.RemoveFromList),

        // 드라이브째 없으면 외장 드라이브를 잠시 뺀 것일 수 있으므로 건너뛰기를 기본값으로 둔다.
        SyncState.DriveMissing => (
            new List<ActionOption>
            {
                new("건너뛰기", SyncAction.Skip),
                new("목록에서 빼기", SyncAction.RemoveFromList)
            },
            SyncAction.Skip),

        _ => (new List<ActionOption> { new("건너뛰기", SyncAction.Skip) }, SyncAction.Skip)
    };

    private static string StatusLabel(SyncState state) => state switch
    {
        SyncState.NoChange => "변경 없음",
        SyncState.RemoteOnlyChanged => "온라인만 변경",
        SyncState.LocalOnlyChanged => "로컬만 변경",
        SyncState.Conflict => "충돌",
        SyncState.LocalMissing => "로컬에 없음",
        SyncState.DriveMissing => "드라이브 연결 안 됨",
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
