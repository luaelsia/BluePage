using Microsoft365OfficeWebLauncher.OneDrive;
using Microsoft365OfficeWebLauncher.UI.Controls;

namespace Microsoft365OfficeWebLauncher.UI;

/// <summary>
/// "동기화 검토" 창. 창을 먼저 띄운 뒤 파일마다 상태를 조회해 채워 넣는다.
/// 파일마다 감지된 상태에 맞는 동작만 고를 수 있고, 기본값은 이미 합리적으로 선택돼 있다.
/// "건너뛰기"를 고른 파일과 필터로 숨긴 파일은 [선택 항목 동기화]를 눌러도 처리되지 않는다.
/// </summary>
public sealed class SyncSelectionDialog : Form
{
    private sealed record ActionOption(string Display, SyncAction Action);

    /// <summary>표 한 줄의 상태. 조회 중이면 Detection이 null이고 Loading이 true.</summary>
    private sealed class RowModel
    {
        public required string Path { get; init; }
        public SyncDetection? Detection { get; set; }
        public bool Loading { get; set; } = true;
        public bool Failed { get; set; }
        public SyncAction Action { get; set; } = SyncAction.Skip;

        public bool Editable => !Loading && !Failed && Detection is not null && OptionsFor(Detection.DetectedState).Options.Count > 1;
    }

    /// <summary>상태 필터 항목. States가 null이면 전체.</summary>
    private sealed record StatusFilter(string Display, SyncState[]? States);

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

    private readonly IReadOnlyList<string> _paths;
    private readonly Func<string, CancellationToken, Task<SyncDetection>> _detect;
    private readonly Action<string, Exception>? _onDetectFailed;
    private readonly CancellationTokenSource _closing = new();
    private readonly DataGridView _grid;
    private SelectBox _statusFilter = null!;
    private TextBox _nameFilter = null!;
    private Label _filterCountLabel = null!;
    private ModernButton _refreshButton = null!;
    private ModernButton _runButton = null!;
    private bool _loading;
    private int _loadedCount;

    /// <summary>필터로 숨긴 행은 포함하지 않는다(화면에 보이는 행만 실행).</summary>
    public IReadOnlyList<(string Path, SyncAction Action)> SelectedActions { get; private set; } = Array.Empty<(string, SyncAction)>();

    public SyncSelectionDialog(
        IReadOnlyList<string> paths,
        Func<string, CancellationToken, Task<SyncDetection>> detect,
        Action<string, Exception>? onDetectFailed = null)
    {
        _paths = paths;
        _detect = detect;
        _onDetectFailed = onDetectFailed;

        Text = "동기화 검토";
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        MaximizeBox = true;
        StartPosition = FormStartPosition.CenterScreen;
        Size = new Size(UiDraw.S(980), UiDraw.S(640));
        MinimumSize = new Size(UiDraw.S(780), UiDraw.S(440));
        Font = new Font("Segoe UI Variable Text", 9.5F);
        Padding = UiDraw.S(18, 14, 18, 14);

        // 표는 둥근 카드 안에 넣는다. Dock 순서상 Fill을 먼저 넣어야 위/아래 패널 사이에 자리 잡는다.
        _grid = BuildGrid();
        var card = new CardPanel { Dock = DockStyle.Fill, Padding = UiDraw.S(10, 8, 10, 8) };
        card.Controls.Add(_grid);
        Controls.Add(card);
        Controls.Add(BuildFilterPanel());
        Controls.Add(BuildButtonPanel());

        foreach (var path in paths)
        {
            var index = _grid.Rows.Add(System.IO.Path.GetFileName(path), string.Empty, null, string.Empty);
            var row = _grid.Rows[index];
            row.Tag = new RowModel { Path = path };
            row.Cells["File"].ToolTipText = path;
        }
        ApplyFilter();

        ThemeApplier.Apply(this, AppTheme.Current);

        Shown += async (_, _) => await LoadAsync();
        FormClosing += (_, _) => _closing.Cancel();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _closing.Dispose();
        }
        base.Dispose(disposing);
    }

    // ---------- 화면 구성 ----------

    private DataGridView BuildGrid()
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            RowHeadersVisible = false,
            ReadOnly = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            ColumnHeadersHeight = UiDraw.S(38),
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            RowTemplate = { Height = UiDraw.S(40) }
        };

        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "File", HeaderText = "파일", FillWeight = 38 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Status", HeaderText = "상태", FillWeight = 16 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "LastSynced", HeaderText = "마지막 동기화", FillWeight = 16 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Action", HeaderText = "실행할 동작", FillWeight = 30 });

        grid.CellFormatting += OnCellFormatting;
        grid.CellPainting += OnCellPainting;
        grid.SortCompare += OnSortCompare;
        grid.CellMouseClick += OnCellMouseClick;
        grid.CellMouseEnter += (_, e) => grid.Cursor = IsEditableActionCell(e.RowIndex, e.ColumnIndex) ? Cursors.Hand : Cursors.Default;
        grid.CellMouseLeave += (_, _) => grid.Cursor = Cursors.Default;
        grid.KeyDown += (_, e) =>
        {
            // 키보드로도 동작을 고를 수 있게 한다(선택한 행에서 Space 또는 Enter).
            if (e.KeyCode is Keys.Space or Keys.Enter && grid.CurrentRow is { } row && IsEditableActionCell(row.Index, grid.Columns["Action"]!.Index))
            {
                ShowActionMenu(row.Index);
                e.Handled = true;
            }
        };
        return grid;
    }

    private Control BuildFilterPanel()
    {
        var bar = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 2, RowCount = 1, AutoSize = true, Padding = new Padding(0, 0, 0, UiDraw.S(10)) };
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var filters = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, WrapContents = false, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0) };
        filters.Controls.Add(new Label { Text = "상태", AutoSize = true, Margin = UiDraw.S(2, 9, 8, 0) });
        _statusFilter = new SelectBox(StatusFilters.Select(filter => filter.Display).ToArray()) { Width = UiDraw.S(250), Margin = UiDraw.S(0, 0, 20, 0) };
        _statusFilter.SelectedIndex = 0;
        _statusFilter.SelectedIndexChanged += (_, _) => ApplyFilter();
        filters.Controls.Add(_statusFilter);

        filters.Controls.Add(new Label { Text = "파일 이름", AutoSize = true, Margin = UiDraw.S(0, 9, 8, 0) });
        _nameFilter = new TextBox { Width = UiDraw.S(200), Margin = UiDraw.S(0, 7, 20, 0), PlaceholderText = "이름 일부 입력" };
        _nameFilter.TextChanged += (_, _) => ApplyFilter();
        filters.Controls.Add(_nameFilter);

        _filterCountLabel = new Label { AutoSize = true, Tag = ThemeApplier.SecondaryTag, Margin = UiDraw.S(0, 9, 0, 0) };
        filters.Controls.Add(_filterCountLabel);
        bar.Controls.Add(filters, 0, 0);

        _refreshButton = new ModernButton { Text = "새로고침", AutoSize = true, Anchor = AnchorStyles.Right, Margin = new Padding(UiDraw.S(12), 0, 0, 0) };
        _refreshButton.Click += async (_, _) => await LoadAsync();
        bar.Controls.Add(_refreshButton, 1, 0);
        return bar;
    }

    private FlowLayoutPanel BuildButtonPanel()
    {
        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Padding = new Padding(0, UiDraw.S(12), 0, 0)
        };

        var cancelButton = new ModernButton { Text = "취소", AutoSize = true, DialogResult = DialogResult.Cancel, Margin = new Padding(UiDraw.S(8), 0, 0, 0) };
        panel.Controls.Add(cancelButton);

        _runButton = new ModernButton { Text = "선택 항목 동기화", AutoSize = true, IsPrimary = true };
        _runButton.Click += (_, _) => OnRunClicked();
        panel.Controls.Add(_runButton);

        AcceptButton = _runButton;
        CancelButton = cancelButton;
        return panel;
    }

    // ---------- 조회 ----------

    /// <summary>
    /// 모든 행을 다시 조회한다. 처음 열 때와 [새로고침]에서 쓴다.
    /// 이미 고른 동작은 상태가 그대로면 유지하고, 상태가 바뀐 행만 기본 동작으로 되돌린다.
    /// </summary>
    private async Task LoadAsync()
    {
        if (_loading)
        {
            return;
        }

        _loading = true;
        _loadedCount = 0;
        _refreshButton.Enabled = false;
        _runButton.Enabled = false;

        var rows = _grid.Rows.Cast<DataGridViewRow>().ToList();
        foreach (var row in rows)
        {
            ((RowModel)row.Tag!).Loading = true;
        }
        _grid.Invalidate();
        UpdateCountLabel();

        try
        {
            foreach (var row in rows)
            {
                var model = (RowModel)row.Tag!;
                var previousState = model.Detection?.DetectedState;
                try
                {
                    var detection = await _detect(model.Path, _closing.Token);
                    model.Detection = detection;
                    model.Failed = false;
                    if (previousState != detection.DetectedState)
                    {
                        model.Action = OptionsFor(detection.DetectedState).Default;
                    }
                }
                catch (OperationCanceledException) when (_closing.IsCancellationRequested)
                {
                    return;
                }
                catch (Exception ex)
                {
                    _onDetectFailed?.Invoke(model.Path, ex);
                    model.Failed = true;
                    model.Action = SyncAction.Skip;
                }
                finally
                {
                    model.Loading = false;
                }

                if (IsDisposed)
                {
                    return;
                }
                _loadedCount++;
                _grid.InvalidateRow(row.Index);
                UpdateCountLabel();
            }
        }
        finally
        {
            _loading = false;
            if (!IsDisposed)
            {
                _refreshButton.Enabled = true;
                _runButton.Enabled = true;
                SortRows();
                ApplyFilter();
            }
        }
    }

    /// <summary>사용자가 열 제목으로 고른 정렬이 있으면 그대로, 없으면 상태 순으로 정렬한다.</summary>
    private void SortRows()
    {
        var column = _grid.SortedColumn ?? _grid.Columns["Status"]!;
        var direction = _grid.SortedColumn is null || _grid.SortOrder != SortOrder.Descending
            ? System.ComponentModel.ListSortDirection.Ascending
            : System.ComponentModel.ListSortDirection.Descending;
        _grid.Sort(column, direction);
    }

    // ---------- 표시 ----------

    private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.RowIndex < 0 || _grid.Rows[e.RowIndex].Tag is not RowModel model)
        {
            return;
        }

        var theme = AppTheme.Current;
        var column = _grid.Columns[e.ColumnIndex].Name;
        switch (column)
        {
            case "Status":
                e.Value = model.Loading ? "확인 중..." : model.Failed ? "확인 실패" : StatusLabel(model.Detection!.DetectedState);
                e.CellStyle!.ForeColor = model.Loading ? theme.TextSecondary
                    : model.Failed ? theme.DangerSoftText
                    : StatusColor(model.Detection!.DetectedState);
                e.FormattingApplied = true;
                break;

            case "LastSynced":
                e.Value = model.Detection is null ? string.Empty : FormatRelative(model.Detection.LastSyncedUtc);
                e.CellStyle!.ForeColor = theme.TextSecondary;
                e.FormattingApplied = true;
                break;

            case "File":
                var dimmed = model.Detection?.DetectedState == SyncState.NoChange;
                e.CellStyle!.ForeColor = dimmed ? theme.TextSecondary : theme.TextPrimary;
                break;
        }
    }

    /// <summary>"실행할 동작" 칸은 앱의 다른 드롭다운과 같은 둥근 칸과 ⌄로 직접 그린다.</summary>
    private void OnCellPainting(object? sender, DataGridViewCellPaintingEventArgs e)
    {
        if (e.RowIndex < 0 || _grid.Columns[e.ColumnIndex].Name != "Action" || _grid.Rows[e.RowIndex].Tag is not RowModel model)
        {
            return;
        }

        var theme = AppTheme.Current;
        e.PaintBackground(e.CellBounds, true);
        var g = e.Graphics!;
        var font = e.CellStyle?.Font ?? _grid.Font;

        var text = model.Loading ? string.Empty : model.Failed ? "건너뛰기" : ActionLabel(model);
        if (!model.Editable)
        {
            var textBounds = Rectangle.Inflate(e.CellBounds, -UiDraw.S(12), 0);
            TextRenderer.DrawText(g, text, font, textBounds, theme.TextSecondary,
                TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
            e.Handled = true;
            return;
        }

        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        var box = new RectangleF(
            e.CellBounds.Left + UiDraw.S(4),
            e.CellBounds.Top + UiDraw.S(5),
            e.CellBounds.Width - UiDraw.S(12),
            e.CellBounds.Height - UiDraw.S(10));
        using (var path = UiDraw.RoundedRect(box, UiDraw.S(9)))
        using (var fill = new SolidBrush(theme.ButtonBackground))
        using (var border = new Pen(theme.ButtonBorder))
        {
            g.FillPath(fill, path);
            g.DrawPath(border, path);
        }

        var inner = Rectangle.Round(box);
        TextRenderer.DrawText(g, text, font, new Rectangle(inner.Left + UiDraw.S(10), inner.Top, inner.Width - UiDraw.S(36), inner.Height),
            theme.TextPrimary, TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
        using var iconFont = UiDraw.IconFont(7.5F);
        TextRenderer.DrawText(g, Glyphs.ChevronDown, iconFont, new Rectangle(inner.Right - UiDraw.S(28), inner.Top, UiDraw.S(22), inner.Height),
            theme.TextSecondary, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        e.Handled = true;
    }

    // ---------- 정렬 ----------

    /// <summary>상태는 중요한 순서(충돌 먼저)로, 마지막 동기화는 실제 시각으로 정렬한다. 같으면 파일 이름순.</summary>
    private void OnSortCompare(object? sender, DataGridViewSortCompareEventArgs e)
    {
        var left = (RowModel)_grid.Rows[e.RowIndex1].Tag!;
        var right = (RowModel)_grid.Rows[e.RowIndex2].Tag!;
        var column = e.Column.Name;

        int result;
        if (column == "Status")
        {
            result = StatusRank(left).CompareTo(StatusRank(right));
        }
        else if (column == "LastSynced")
        {
            // 최근 것이 위로 오도록 시각을 거꾸로 비교한다(오름차순 = 최근 순).
            var leftTime = left.Detection?.LastSyncedUtc ?? DateTimeOffset.MinValue;
            var rightTime = right.Detection?.LastSyncedUtc ?? DateTimeOffset.MinValue;
            result = rightTime.CompareTo(leftTime);
        }
        else if (column == "Action")
        {
            result = string.Compare(ActionLabel(left), ActionLabel(right), StringComparison.CurrentCulture);
        }
        else
        {
            result = 0;
        }

        if (result == 0)
        {
            result = string.Compare(System.IO.Path.GetFileName(left.Path), System.IO.Path.GetFileName(right.Path), StringComparison.CurrentCultureIgnoreCase);
        }
        e.SortResult = result;
        e.Handled = true;
    }

    private static int StatusRank(RowModel model)
    {
        if (model.Loading) return 90;
        if (model.Failed) return 80;
        return model.Detection!.DetectedState switch
        {
            SyncState.Conflict => 0,
            SyncState.RemoteOnlyChanged => 1,
            SyncState.LocalOnlyChanged => 2,
            SyncState.LocalMissing => 3,
            SyncState.DriveMissing => 4,
            SyncState.NoChange => 5,
            _ => 6
        };
    }

    // ---------- 동작 고르기 ----------

    private bool IsEditableActionCell(int rowIndex, int columnIndex) =>
        rowIndex >= 0 && columnIndex >= 0 &&
        _grid.Columns[columnIndex].Name == "Action" &&
        _grid.Rows[rowIndex].Tag is RowModel { Editable: true };

    private void OnCellMouseClick(object? sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left && IsEditableActionCell(e.RowIndex, e.ColumnIndex))
        {
            ShowActionMenu(e.RowIndex);
        }
    }

    private void ShowActionMenu(int rowIndex)
    {
        var model = (RowModel)_grid.Rows[rowIndex].Tag!;
        var (options, _) = OptionsFor(model.Detection!.DetectedState);

        var menu = new ContextMenuStrip { ShowImageMargin = false, ShowCheckMargin = true, Font = _grid.Font };
        foreach (var option in options)
        {
            var item = new ToolStripMenuItem(option.Display) { Checked = option.Action == model.Action, Padding = new Padding(0, UiDraw.S(3), 0, UiDraw.S(3)) };
            item.Click += (_, _) =>
            {
                model.Action = option.Action;
                _grid.InvalidateRow(rowIndex);
            };
            menu.Items.Add(item);
        }
        ThemedMenu.Apply(menu);

        var cell = _grid.GetCellDisplayRectangle(_grid.Columns["Action"]!.Index, rowIndex, cutOverflow: false);
        menu.MinimumSize = new Size(cell.Width - UiDraw.S(12), 0);
        menu.Closed += (_, _) => BeginInvoke(menu.Dispose);
        menu.Show(_grid, new Point(cell.Left + UiDraw.S(4), cell.Bottom - UiDraw.S(4)));
    }

    // ---------- 필터와 실행 ----------

    private void ApplyFilter()
    {
        var states = _statusFilter.SelectedIndex >= 0 ? StatusFilters[_statusFilter.SelectedIndex].States : null;
        var keyword = _nameFilter.Text.Trim();

        // 현재 셀이 있는 행은 숨길 수 없으므로 먼저 현재 셀을 해제한다.
        _grid.CurrentCell = null;

        foreach (DataGridViewRow row in _grid.Rows)
        {
            var model = (RowModel)row.Tag!;
            var stateMatches = states is null ||
                               (model.Detection is not null && !model.Loading && states.Contains(model.Detection.DetectedState));
            var nameMatches = keyword.Length == 0 ||
                              System.IO.Path.GetFileName(model.Path).Contains(keyword, StringComparison.CurrentCultureIgnoreCase);
            row.Visible = stateMatches && nameMatches;
        }
        UpdateCountLabel();
    }

    private void UpdateCountLabel()
    {
        var total = _grid.Rows.Count;
        if (_loading)
        {
            _filterCountLabel.Text = $"확인 중 {_loadedCount}/{total}";
            return;
        }

        var visible = _grid.Rows.Cast<DataGridViewRow>().Count(row => row.Visible);
        _filterCountLabel.Text = visible == total
            ? $"{total}개"
            : $"{total}개 중 {visible}개 표시 (숨긴 파일은 실행하지 않음)";
    }

    private void OnRunClicked()
    {
        if (_loading)
        {
            return;
        }

        var selected = new List<(string Path, SyncAction Action)>();
        foreach (DataGridViewRow row in _grid.Rows)
        {
            var model = (RowModel)row.Tag!;
            if (!row.Visible || model.Loading || model.Failed || model.Action == SyncAction.Skip)
            {
                continue;
            }
            selected.Add((model.Path, model.Action));
        }

        SelectedActions = selected;
        DialogResult = DialogResult.OK;
        Close();
    }

    // ---------- 상태별 선택지와 문구 ----------

    private static string ActionLabel(RowModel model)
    {
        if (model.Detection is null)
        {
            return string.Empty;
        }
        var option = OptionsFor(model.Detection.DetectedState).Options.FirstOrDefault(o => o.Action == model.Action);
        return option?.Display ?? "건너뛰기";
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

    private static Color StatusColor(SyncState state)
    {
        var theme = AppTheme.Current;
        return state switch
        {
            SyncState.Conflict => theme.DangerSoftText,
            SyncState.RemoteOnlyChanged or SyncState.LocalOnlyChanged => theme.AccentSoftText,
            SyncState.LocalMissing or SyncState.DriveMissing => theme.WarningSoftText,
            _ => theme.TextSecondary
        };
    }

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
