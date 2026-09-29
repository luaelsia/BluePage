using System.Diagnostics;
using Microsoft365OfficeWebLauncher.Auth;
using Microsoft365OfficeWebLauncher.Cloud;
using Microsoft365OfficeWebLauncher.Config;
using Microsoft365OfficeWebLauncher.Logging;
using Microsoft365OfficeWebLauncher.OneDrive;
using Microsoft365OfficeWebLauncher.UI;

namespace Microsoft365OfficeWebLauncher.Core;

/// <summary>확장자 판별 → 동기화(업로드/다운로드) → 기본 브라우저로 webUrl 열기를 조율한다.</summary>
public sealed class LaunchOrchestrator
{
    private readonly DocumentTypeCatalog _catalog;
    private readonly SyncCoordinator _syncCoordinator;
    private readonly UploadManifest _manifest;
    private readonly FileLogger _logger;
    private readonly DeferredSyncRegistry _deferredSyncRegistry;
    private readonly AppConfig _config;
    private readonly HashSet<string> _reportedMissingFiles = new(StringComparer.OrdinalIgnoreCase);

    public LaunchOrchestrator(
        DocumentTypeCatalog catalog,
        SyncCoordinator syncCoordinator,
        UploadManifest manifest,
        FileLogger logger,
        AppConfig config)
    {
        _catalog = catalog;
        _syncCoordinator = syncCoordinator;
        _manifest = manifest;
        _logger = logger;
        _config = config;
        _deferredSyncRegistry = new DeferredSyncRegistry();
    }

    public async Task<int> OpenAsync(string filePath, CancellationToken ct)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath))
        {
            _logger.Error($"파일을 찾을 수 없습니다: {fullPath}");
            ShowError($"파일을 찾을 수 없습니다:\n{fullPath}");
            return 1;
        }

        if (!_catalog.TryResolve(fullPath, out var appDefinition))
        {
            _logger.Error($"지원하지 않는 확장자입니다: {Path.GetExtension(fullPath)}");
            ShowError($"지원하지 않는 파일 형식입니다: {Path.GetExtension(fullPath)}\n" +
                      "appsettings.json의 documentTypes에 추가하면 지원할 수 있습니다.");
            return 1;
        }

        _logger.Info($"열기 시작: {fullPath} ({appDefinition.OfficeApp})");
        _deferredSyncRegistry.Resume(fullPath);

        if (string.Equals(Path.GetExtension(fullPath), ".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            ShowWarning("XLSM 파일을 웹에서 열 수 있지만 매크로는 실행되지 않습니다.");
        }

        var provider = SelectProvider(fullPath, appDefinition);
        if (provider is null)
        {
            _logger.Info($"사용자가 웹 Office 선택을 취소했습니다: {fullPath}");
            return 0;
        }

        SyncResult result;
        try
        {
            result = await _syncCoordinator.PrepareForOpenAsync(fullPath, ct, provider);
        }
        catch (Exception ex) when (GraphErrorHelper.IsResourceLocked(ex))
        {
            _logger.Warn($"Office Web 편집 잠금 감지, 문서 닫힘 대기 시작: {fullPath}");

            using var waitDialog = new WebDocumentUnlockDialog(
                fullPath,
                checkCt => _syncCoordinator.GetRemoteLockStateAsync(fullPath, checkCt),
                retryCt => _syncCoordinator.PrepareForOpenAsync(fullPath, retryCt, provider));
            var dialogResult = waitDialog.ShowDialog();

            if (dialogResult == DialogResult.Cancel ||
                (waitDialog.SyncResult is null && waitDialog.SyncError is null))
            {
                if (waitDialog.SynchronizationStopped)
                {
                    _deferredSyncRegistry.Defer(fullPath);
                }
                _logger.Info($"사용자가 웹 문서 닫힘 대기를 취소했습니다: {fullPath}");
                return 0;
            }

            if (waitDialog.SyncError is not null)
            {
                _logger.Error($"웹 문서 잠금 해제 후 동기화 실패: {fullPath}", waitDialog.SyncError);
                ShowError($"문서가 닫힌 후 동기화하는 중 오류가 발생했습니다.\n\n{waitDialog.SyncError.Message}");
                return 1;
            }

            result = waitDialog.SyncResult!;
            _logger.Info($"웹 문서 잠금 해제 후 동기화 완료: {fullPath} (상태: {result.State})");
        }
        catch (GoogleReauthRequiredException ex)
        {
            _logger.Warn($"Google 로그인이 만료되어 동기화하지 못했습니다: {fullPath}");
            ShowError(ex.Message);
            return 1;
        }
        catch (Exception ex)
        {
            _logger.Error($"동기화 실패: {fullPath}", ex);

            // 이전에 이미 업로드된 적 있는 파일이면, 반영은 실패했더라도 알고 있는 온라인 사본을 그대로 열어준다
            // (동기화 실패 = 문서를 아예 못 여는 것보다는 낫다).
            var existingEntry = _manifest.Get(fullPath);
            if (existingEntry is not null && !string.IsNullOrEmpty(existingEntry.WebUrl))
            {
                ShowInfo("온라인 동기화에 실패했습니다(자세한 내용은 로그 참고).\n" +
                         "일단 온라인의 최신 버전을 그대로 엽니다.");

                OpenInBrowser(existingEntry.WebUrl);
                _logger.Info($"동기화 실패 후 기존 온라인 사본을 열었습니다: {existingEntry.WebUrl}");
                return 0;
            }

            ShowError($"동기화에 실패해 문서를 열 수 없습니다.\n\n{ex.Message}");
            return 1;
        }

        if (result.State == SyncState.Skipped)
        {
            _logger.Info($"사용자 선택으로 문서를 열지 않고 동기화를 종료합니다: {fullPath}");
            return 0;
        }

        if (result.State == SyncState.Conflict)
        {
            ShowConflictCopySaved(result.ConflictCopyPath);
        }

        OpenInBrowser(result.WebUrl);
        _logger.Info($"브라우저로 열기 완료: {result.WebUrl} (상태: {result.State})");
        return 0;
    }

    /// <summary>
    /// --url 흐름: 창을 하나도 띄우지 않고 동기화한 뒤 보기용 주소를 돌려준다(브라우저는 열지 않는다).
    /// AI가 실행하는 명령이라 로그인 창, 충돌 창, 잠금 대기 창이 뜨면 아무도 누르지 않은 채 멈추므로,
    /// 그런 상황은 모두 오류 코드로 끝낸다. SyncCoordinator는 NonInteractiveConflictResolver로 만들어야 한다.
    /// </summary>
    public async Task<AiUrlResult> ResolveUrlForAiAsync(string filePath, CancellationToken ct)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath))
        {
            return AiUrlResult.Fail(fullPath, AiUrlError.NotFound, "파일을 찾을 수 없습니다.");
        }

        if (!_catalog.TryResolve(fullPath, out _))
        {
            return AiUrlResult.Fail(fullPath, AiUrlError.Unsupported, $"지원하지 않는 파일 형식입니다: {Path.GetExtension(fullPath)}");
        }

        var entry = _manifest.Get(fullPath);
        if (!AiAccess.IsRegistered(entry))
        {
            return AiUrlResult.Fail(fullPath, AiUrlError.NotRegistered,
                "BluePage로 한 번도 연 적 없는 문서입니다. 사용자가 탐색기에서 이 파일을 BluePage로 한 번 열어야 합니다.");
        }

        try
        {
            // 공급자를 넘기지 않으면 매니페스트에 기록된 공급자를 그대로 쓴다(선택 창이 뜨지 않는다).
            var result = await _syncCoordinator.PrepareForOpenAsync(fullPath, ct);
            _logger.Info($"AI 주소 요청 처리: {fullPath} (상태: {result.State})");

            if (result.State is SyncState.Skipped or SyncState.Conflict)
            {
                return AiUrlResult.Fail(fullPath, AiUrlError.Conflict,
                    "로컬 파일과 온라인 사본이 모두 바뀌어 동기화하지 않았습니다. 사용자가 BluePage로 문서를 열어 충돌을 직접 처리해야 합니다.");
            }

            return AiUrlResult.Success(fullPath, _manifest.Get(fullPath)?.Provider ?? entry!.Provider, result.State, result.WebUrl);
        }
        catch (Exception ex) when (GraphErrorHelper.IsResourceLocked(ex))
        {
            _logger.Warn($"AI 주소 요청: Office Web 편집 잠금으로 동기화하지 못했습니다: {fullPath}");
            return AiUrlResult.Fail(fullPath, AiUrlError.Locked,
                "온라인 문서가 웹에서 편집 중이라 잠겨 있습니다. 열려 있는 웹 문서 탭을 모두 닫은 뒤 다시 실행하세요.");
        }
        catch (Exception ex) when (FindInChain<GoogleReauthRequiredException>(ex) is not null ||
                                   FindInChain<MicrosoftSignInRequiredException>(ex) is not null)
        {
            _logger.Warn($"AI 주소 요청: 로그인이 필요해 동기화하지 못했습니다: {fullPath}");
            var authError = (Exception?)FindInChain<GoogleReauthRequiredException>(ex) ?? FindInChain<MicrosoftSignInRequiredException>(ex)!;
            return AiUrlResult.Fail(fullPath, AiUrlError.AuthRequired, authError.Message);
        }
        catch (Exception ex)
        {
            _logger.Error($"AI 주소 요청 처리 실패: {fullPath}", ex);
            return AiUrlResult.Fail(fullPath, AiUrlError.Failed, $"동기화에 실패했습니다: {ex.Message}");
        }
    }

    private static T? FindInChain<T>(Exception? ex) where T : Exception
    {
        while (ex is not null)
        {
            if (ex is T match)
            {
                return match;
            }

            if (ex is AggregateException aggregate)
            {
                return aggregate.Flatten().InnerExceptions.Select(FindInChain<T>).FirstOrDefault(found => found is not null);
            }

            ex = ex.InnerException;
        }

        return null;
    }

    public async Task<int> SyncOneAsync(string filePath, CancellationToken ct)
    {
        var fullPath = Path.GetFullPath(filePath);
        _deferredSyncRegistry.Resume(fullPath);
        if (!File.Exists(fullPath))
        {
            _logger.Error($"동기화 대상 파일을 찾을 수 없습니다: {fullPath}");
            ShowError($"파일을 찾을 수 없습니다:\n{fullPath}");
            return 1;
        }

        var result = await _syncCoordinator.PrepareForOpenAsync(fullPath, ct);
        _logger.Info($"동기화 완료: {fullPath} (상태: {result.State})");

        if (result.State == SyncState.Conflict)
        {
            ShowConflictCopySaved(result.ConflictCopyPath);
        }

        return 0;
    }

    public async Task<int> SyncAllAsync(CancellationToken ct)
    {
        var paths = _manifest.AllEntries.Keys.ToList();
        _logger.Debug($"전체 동기화 시작: {paths.Count}개 파일");

        // 백그라운드 주기마다 호출되므로 아무 일도 없는 회차는 Info 로그를 남기지 않는다.
        // 반영이나 실패가 있었던 회차만 끝에 요약 한 줄을 남긴다.
        var appliedCount = 0;
        var failedCount = 0;

        // Google 로그인이 만료된 상태에서 남은 파일까지 계속 시도하면 매 주기마다 같은 실패가 수십 번 쌓인다.
        // 한 번 확인되면 이번 회차의 나머지 Google 파일은 건너뛰고, 사용자가 다시 로그인한 뒤에 재개한다.
        var googleSignInExpired = false;

        foreach (var path in paths)
        {
            if (googleSignInExpired && IsGoogleEntry(path))
            {
                continue;
            }

            if (_deferredSyncRegistry.IsDeferred(path))
            {
                _logger.Debug($"사용자가 동기화하지 않기로 한 파일은 백그라운드에서 건너뜁니다: {path}");
                continue;
            }

            if (!File.Exists(path))
            {
                _deferredSyncRegistry.Resume(path);
                // 같은 파일은 사라진 것을 처음 발견했을 때 한 번만 경고한다(다시 생겼다가 사라지면 다시 경고).
                if (_reportedMissingFiles.Add(path))
                {
                    _logger.Warn($"로컬에서 사라진 파일은 건너뜁니다: {path}");
                }
                continue;
            }
            _reportedMissingFiles.Remove(path);

            try
            {
                var result = await _syncCoordinator.PrepareForOpenAsync(path, ct);
                if (result.State == SyncState.NoChange)
                {
                    _logger.Debug($"동기화됨: {path} (상태: {result.State})");
                }
                else
                {
                    _logger.Info($"동기화됨: {path} (상태: {result.State})");
                    appliedCount++;
                }
                if (result.State == SyncState.Skipped)
                {
                    _deferredSyncRegistry.Defer(path);
                }
            }
            catch (Exception ex) when (GraphErrorHelper.IsResourceLocked(ex))
            {
                _logger.Warn($"백그라운드 동기화 중 Office Web 편집 잠금 감지, 문서 닫힘 대기 시작: {path}");

                using var waitDialog = new WebDocumentUnlockDialog(
                    path,
                    checkCt => _syncCoordinator.GetRemoteLockStateAsync(path, checkCt),
                    retryCt => _syncCoordinator.PrepareForOpenAsync(path, retryCt));
                var dialogResult = waitDialog.ShowDialog();

                if (dialogResult == DialogResult.Cancel ||
                    (waitDialog.SyncResult is null && waitDialog.SyncError is null))
                {
                    if (waitDialog.SynchronizationStopped)
                    {
                        _deferredSyncRegistry.Defer(path);
                    }
                    _logger.Info($"사용자가 백그라운드 동기화의 웹 문서 닫힘 대기를 취소했습니다: {path}");
                    continue;
                }

                if (waitDialog.SyncError is not null)
                {
                    _logger.Error($"웹 문서 잠금 해제 후 백그라운드 동기화 실패: {path}", waitDialog.SyncError);
                    failedCount++;
                    ShowError($"문서가 닫힌 후 동기화하는 중 오류가 발생했습니다.\n\n{waitDialog.SyncError.Message}");
                    continue;
                }

                var retryResult = waitDialog.SyncResult!;
                _logger.Info($"웹 문서 잠금 해제 후 백그라운드 동기화 완료: {path} (상태: {retryResult.State})");
                if (retryResult.State != SyncState.NoChange)
                {
                    appliedCount++;
                }

                if (retryResult.State == SyncState.Skipped)
                {
                    _deferredSyncRegistry.Defer(path);
                    continue;
                }

                if (retryResult.State == SyncState.Conflict)
                {
                    ShowConflictCopySaved(retryResult.ConflictCopyPath);
                }

                OpenInBrowser(retryResult.WebUrl);
                _logger.Info($"잠금 해제 후 웹 문서를 다시 열었습니다: {retryResult.WebUrl}");
                ShowInfo($"동기화를 완료하고 웹 문서를 다시 열었습니다:\n{Path.GetFileName(path)}");
            }
            catch (GoogleReauthRequiredException)
            {
                googleSignInExpired = true;
                failedCount++;
                _logger.Warn("Google 로그인이 만료되어 이번 회차의 Google 문서 동기화를 건너뜁니다. " +
                             "BluePage 창에서 Google 계정에 다시 로그인해 주세요.");
            }
            catch (Exception ex)
            {
                _logger.Error($"동기화 실패: {path}", ex);
                failedCount++;
            }
        }

        if (appliedCount > 0 || failedCount > 0)
        {
            _logger.Info($"전체 동기화 완료: {paths.Count}개 확인, 반영 {appliedCount}, 실패 {failedCount}");
        }

        return 0;
    }

    private bool IsGoogleEntry(string path) =>
        _manifest.Get(path) is { } entry &&
        CloudProviderNames.TryParse(entry.Provider, out var provider) &&
        provider == CloudProvider.Google;

    /// <summary>동기화 검토 창용: 쓰기 없이 상태만 조회한다.</summary>
    public Task<SyncDetection> DetectSyncStatusAsync(string filePath, CancellationToken ct) =>
        _syncCoordinator.DetectAsync(filePath, ct);

    /// <summary>동기화 검토 창용: 사용자가 고른 동작 하나만 실행한다.</summary>
    public Task<SyncResult> ApplySyncActionAsync(string filePath, SyncAction action, CancellationToken ct)
    {
        _deferredSyncRegistry.Resume(Path.GetFullPath(filePath));

        if (action == SyncAction.RemoveFromList)
        {
            // 동기화 목록에서만 뺀다. OneDrive/Google Drive의 온라인 사본과 로컬 백업은 그대로 둔다.
            _manifest.Remove(filePath);
            _manifest.Save();
            _reportedMissingFiles.Remove(filePath);
            _logger.Info($"동기화 목록에서 제거: {filePath}");
            return Task.FromResult(new SyncResult(SyncState.Skipped, string.Empty, null));
        }

        return _syncCoordinator.ApplyAsync(filePath, action, ct);
    }

    /// <summary>GUI 프로세스에서만 실제 토스트 알림을 연결한다(헤드리스 오픈/CLI는 호출하지 않음).</summary>
    public void AttachActivityReporter(ISyncActivityReporter reporter) => _syncCoordinator.AttachActivityReporter(reporter);

    private CloudProvider? SelectProvider(string filePath, OfficeAppDefinition definition)
    {
        if (definition.SupportedProviders.Count == 1)
        {
            return definition.SupportedProviders.Single();
        }

        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        var preference = _config.DocumentProviderPreferences.TryGetValue(extension, out var perExtension) &&
                         !string.Equals(perExtension, "Default", StringComparison.OrdinalIgnoreCase)
            ? perExtension
            : _config.PreferredCloudProvider;

        if (CloudProviderNames.TryParse(preference, out var configured) && definition.Supports(configured))
        {
            return configured;
        }

        using var dialog = new CloudProviderDialog(filePath, definition.SupportedProviders);
        if (dialog.ShowDialog() != DialogResult.OK || dialog.SelectedProvider is null)
        {
            return null;
        }

        if (dialog.RememberAsDefault)
        {
            _config.PreferredCloudProvider = dialog.SelectedProvider.Value.ToString();
            ConfigLoader.Save(_config);
        }
        return dialog.SelectedProvider;
    }

    private static void OpenInBrowser(string url)
    {
        Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
    }

    /// <summary>충돌 사본 저장 안내. [폴더 열기]를 누르면 사본 파일이 선택된 상태로 탐색기를 연다.</summary>
    public static void ShowConflictCopySaved(string? conflictCopyPath, IWin32Window? owner = null)
    {
        if (string.IsNullOrEmpty(conflictCopyPath))
        {
            return;
        }

        AppMessageDialog.ShowWithAction(
            owner,
            "로컬 파일과 Office Web의 온라인 사본이 모두 변경되어 자동으로 병합할 수 없습니다.\n" +
            $"온라인 사본을 아래 경로에 별도로 저장했습니다. 두 파일을 확인 후 직접 병합해 주세요.\n\n{conflictCopyPath}",
            "폴더 열기",
            () => ExplorerLauncher.RevealFile(conflictCopyPath));
    }

    private static void ShowError(string message) =>
        AppMessageDialog.Show(message, AppBrand.Name, AppMessageKind.Error);

    private static void ShowInfo(string message) =>
        AppMessageDialog.Show(message, AppBrand.Name, AppMessageKind.Information);

    private static void ShowWarning(string message) =>
        AppMessageDialog.Show(message, AppBrand.Name, AppMessageKind.Warning);
}
