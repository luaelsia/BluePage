using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;
using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft365OfficeWebLauncher.Auth;
using Microsoft365OfficeWebLauncher.Cloud;
using Microsoft365OfficeWebLauncher.Logging;

namespace Microsoft365OfficeWebLauncher.OneDrive;

/// <summary>
/// Microsoft Graph API로 OneDrive의 앱 전용 폴더(App Folder, /me/drive/special/approot)에 업로드/다운로드한다.
/// App Folder를 사용하면 Files.ReadWrite.AppFolder 권한만으로 동작해 최소 권한 원칙을 지킬 수 있다.
///
/// 주의: Files.ReadWrite.AppFolder 권한은 앱 폴더 바깥으로 나가는 호출(GET /me/drive 등 드라이브 전체 메타데이터 조회)을
/// 허용하지 않고 Access Denied를 반환한다. Graph SDK v5의 Me.Drive 빌더는 Items/Special 하위 경로를 노출하지 않으므로,
/// 드라이브 Id는 /me/drive를 거치지 않고 /me/drive/special/approot 응답의 ParentReference.DriveId에서 직접 얻는다
/// (RequestAdapter로 저수준 호출).
/// </summary>
public sealed class OneDriveUploadService : ICloudDriveService
{
    private const long SimpleUploadMaxBytes = 4L * 1024 * 1024; // Graph 문서 기준 단순 업로드 한도
    private const int UploadSliceSize = 320 * 1024; // 320KiB의 배수여야 함(Graph 요구사항)

    private readonly GraphServiceClient _graphClient;
    private readonly GraphAuthService _authService;
    private readonly HttpClient _httpClient = new();
    private readonly FileLogger _logger;
    private string? _driveId;
    private string? _appRootId;
    private string? _appRootWebUrl;
    private string? _bluePageFolderId;
    private string? _bluePageFolderWebUrl;

    public CloudProvider Provider => CloudProvider.Microsoft;
    public string DisplayName => "Microsoft 365";

    public OneDriveUploadService(GraphAuthService authService, FileLogger logger)
    {
        _authService = authService;
        var authProvider = new BaseBearerTokenAuthenticationProvider(new MsalAccessTokenProvider(authService));
        _graphClient = new GraphServiceClient(authProvider);
        _logger = logger;
    }

    /// <summary>
    /// Graph beta의 file.lockInfo를 조회해 실제 업로드 없이 웹 편집 잠금 상태를 확인한다.
    /// 일부 계정/테넌트에서 beta 필드를 제공하지 않으면 Unknown을 반환하고 실제 업로드 때 423으로 다시 검증한다.
    /// </summary>
    public async Task<RemoteLockState> GetLockStateAsync(string driveItemId, CancellationToken ct)
    {
        try
        {
            var (driveId, _) = await GetAppFolderReferenceAsync(ct);
            var authResult = await _authService.AcquireTokenAsync(ct);
            var url = $"https://graph.microsoft.com/beta/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(driveItemId)}?$select=file";
            using var request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
            using var response = await _httpClient.SendAsync(request, ct);

            if (!response.IsSuccessStatusCode)
            {
                _logger.Debug($"웹 잠금 상태 조회 미지원/실패: {(int)response.StatusCode} {response.ReasonPhrase}");
                return RemoteLockState.Unknown;
            }

            await using var content = await response.Content.ReadAsStreamAsync(ct);
            using var json = await JsonDocument.ParseAsync(content, cancellationToken: ct);
            if (!json.RootElement.TryGetProperty("file", out var file) ||
                !file.TryGetProperty("lockInfo", out var lockInfo) ||
                !lockInfo.TryGetProperty("lockType", out var lockTypeElement))
            {
                return RemoteLockState.Unlocked;
            }

            var lockType = lockTypeElement.GetString();
            return string.IsNullOrWhiteSpace(lockType) || string.Equals(lockType, "none", StringComparison.OrdinalIgnoreCase)
                ? RemoteLockState.Unlocked
                : RemoteLockState.Locked;
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.Debug($"웹 잠금 상태 조회 실패: {ex.Message}");
            return RemoteLockState.Unknown;
        }
    }

    public async Task<DriveItem> CreateInAppFolderAsync(string localFilePath, CancellationToken ct)
    {
        var (driveId, appRootId) = await GetBluePageFolderReferenceAsync(ct);
        var fileName = Path.GetFileName(localFilePath);
        var fileInfo = new FileInfo(localFilePath);

        _logger.Info($"OneDrive App Folder/BluePage에 새로 업로드: {fileName} ({fileInfo.Length:N0} bytes)");

        if (fileInfo.Length <= SimpleUploadMaxBytes)
        {
            await using var stream = File.OpenRead(localFilePath);
            var item = await _graphClient.Drives[driveId].Items[appRootId]
                .ItemWithPath(fileName)
                .Content
                .PutAsync(stream, cancellationToken: ct);
            return item ?? throw new InvalidOperationException("업로드 응답이 비어 있습니다.");
        }

        return await UploadLargeFileAsync(
            createSession: sessionBody => _graphClient.Drives[driveId].Items[appRootId]
                .ItemWithPath(fileName).CreateUploadSession.PostAsync(sessionBody, cancellationToken: ct),
            localFilePath,
            ct);
    }

    public async Task<DriveItem> UpdateContentAsync(string driveItemId, string localFilePath, CancellationToken ct)
    {
        var (driveId, _) = await GetAppFolderReferenceAsync(ct);
        var fileInfo = new FileInfo(localFilePath);
        _logger.Info($"기존 OneDrive 항목 갱신: {driveItemId} ({fileInfo.Length:N0} bytes)");

        if (fileInfo.Length <= SimpleUploadMaxBytes)
        {
            await using var stream = File.OpenRead(localFilePath);
            var item = await _graphClient.Drives[driveId].Items[driveItemId].Content.PutAsync(stream, cancellationToken: ct);
            return item ?? throw new InvalidOperationException("업로드 응답이 비어 있습니다.");
        }

        return await UploadLargeFileAsync(
            createSession: sessionBody => _graphClient.Drives[driveId].Items[driveItemId].CreateUploadSession.PostAsync(sessionBody, cancellationToken: ct),
            localFilePath,
            ct);
    }

    public async Task<DriveItem> GetMetadataAsync(string driveItemId, CancellationToken ct)
    {
        var (driveId, bluePageFolderId) = await GetBluePageFolderReferenceAsync(ct);
        var item = await _graphClient.Drives[driveId].Items[driveItemId].GetAsync(cancellationToken: ct);
        if (item is null)
        {
            throw new InvalidOperationException($"드라이브 항목을 찾을 수 없습니다: {driveItemId}");
        }

        // 이전 버전이 App Folder 루트에 올린 파일도 최초 조회 시 BluePage 하위 폴더로 이동한다.
        if (!string.Equals(item.ParentReference?.Id, bluePageFolderId, StringComparison.OrdinalIgnoreCase))
        {
            item = await _graphClient.Drives[driveId].Items[driveItemId].PatchAsync(
                new DriveItem { ParentReference = new ItemReference { Id = bluePageFolderId } },
                cancellationToken: ct) ?? item;
            _logger.Info($"기존 OneDrive 항목을 App Folder/BluePage로 이동: {driveItemId}");
        }
        return item;
    }

    public async Task DownloadContentAsync(string driveItemId, string destinationPath, CancellationToken ct)
    {
        var (driveId, _) = await GetAppFolderReferenceAsync(ct);
        var stream = await _graphClient.Drives[driveId].Items[driveItemId].Content.GetAsync(cancellationToken: ct)
            ?? throw new InvalidOperationException($"드라이브 항목 콘텐츠를 가져올 수 없습니다: {driveItemId}");

        await using (stream)
        {
            var directory = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            await using var fileStream = File.Create(destinationPath);
            await stream.CopyToAsync(fileStream, ct);
        }
    }

    private async Task<DriveItem> UploadLargeFileAsync(
        Func<CreateUploadSessionPostRequestBody, Task<UploadSession?>> createSession,
        string localFilePath,
        CancellationToken ct)
    {
        var sessionBody = new CreateUploadSessionPostRequestBody
        {
            Item = new DriveItemUploadableProperties
            {
                AdditionalData = new Dictionary<string, object>
                {
                    ["@microsoft.graph.conflictBehavior"] = "replace"
                }
            }
        };

        var uploadSession = await createSession(sessionBody)
            ?? throw new InvalidOperationException("업로드 세션 생성에 실패했습니다.");

        await using var fileStream = File.OpenRead(localFilePath);
        var uploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, UploadSliceSize, _graphClient.RequestAdapter);

        var progress = new Progress<long>(uploaded =>
            _logger.Debug($"업로드 진행: {uploaded:N0}/{fileStream.Length:N0} bytes"));

        try
        {
            var result = await uploadTask.UploadAsync(progress, cancellationToken: ct);
            if (!result.UploadSucceeded || result.ItemResponse is null)
            {
                throw new InvalidOperationException("대용량 파일 업로드에 실패했습니다.");
            }

            return result.ItemResponse;
        }
        catch
        {
            // 실패한 업로드 세션을 그대로 두면 서버에 "이 파일명은 업로드 중"이라는 상태가 남아
            // 같은 파일을 다시 열 때마다 "같은 이름의 파일이 업로드 중입니다" 오류로 재시도 자체가 막힌다.
            // 실패 시 세션을 정리해 다음 시도가 걸리지 않게 한다(정리 자체가 실패해도 원래 오류는 그대로 알린다).
            try
            {
                await uploadTask.DeleteSessionAsync(ct);
            }
            catch (Exception cleanupEx)
            {
                _logger.Warn($"업로드 세션 정리 실패: {cleanupEx.Message}");
            }

            throw;
        }
    }

    /// <summary>
    /// GET /me/drive/special/approot 를 저수준으로 호출해 App Folder의 driveItemId와 소속 드라이브 Id를 한 번에 얻는다.
    /// Files.ReadWrite.AppFolder 권한 범위 안에서 허용되는 유일한 진입점이라 /me/drive는 절대 호출하지 않는다.
    /// </summary>
    private async Task<(string DriveId, string AppRootId)> GetAppFolderReferenceAsync(CancellationToken ct)
    {
        if (_driveId is not null && _appRootId is not null)
        {
            return (_driveId, _appRootId);
        }

        var requestInfo = new RequestInformation
        {
            HttpMethod = Method.GET,
            UrlTemplate = "{+baseurl}/me/drive/special/approot",
            PathParameters = new Dictionary<string, object>
            {
                ["baseurl"] = _graphClient.RequestAdapter.BaseUrl!
            }
        };

        var appRoot = await _graphClient.RequestAdapter.SendAsync(requestInfo, DriveItem.CreateFromDiscriminatorValue, cancellationToken: ct)
            ?? throw new InvalidOperationException("OneDrive App Folder를 확인할 수 없습니다.");

        _appRootId = appRoot.Id ?? throw new InvalidOperationException("App Folder의 driveItem Id가 비어 있습니다.");
        _driveId = appRoot.ParentReference?.DriveId ?? throw new InvalidOperationException("App Folder의 드라이브 Id를 확인할 수 없습니다.");
        _appRootWebUrl = appRoot.WebUrl;

        return (_driveId, _appRootId);
    }

    private async Task<(string DriveId, string FolderId)> GetBluePageFolderReferenceAsync(CancellationToken ct)
    {
        var (driveId, appRootId) = await GetAppFolderReferenceAsync(ct);
        if (_bluePageFolderId is not null)
        {
            return (driveId, _bluePageFolderId);
        }

        try
        {
            var existing = await _graphClient.Drives[driveId].Items[appRootId]
                .ItemWithPath("BluePage").GetAsync(cancellationToken: ct);
            if (existing?.Id is not null)
            {
                _bluePageFolderId = existing.Id;
                _bluePageFolderWebUrl = existing.WebUrl;
                return (driveId, _bluePageFolderId);
            }
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            // 폴더가 없으면 바로 아래에서 생성한다.
        }

        var created = await _graphClient.Drives[driveId].Items[appRootId].Children.PostAsync(
            new DriveItem
            {
                Name = "BluePage",
                Folder = new Folder(),
                AdditionalData = new Dictionary<string, object>
                {
                    ["@microsoft.graph.conflictBehavior"] = "fail"
                }
            },
            cancellationToken: ct)
            ?? throw new InvalidOperationException("OneDrive App Folder 안에 BluePage 폴더를 만들지 못했습니다.");

        _bluePageFolderId = created.Id ?? throw new InvalidOperationException("BluePage 폴더 ID가 비어 있습니다.");
        _bluePageFolderWebUrl = created.WebUrl;
        _logger.Info("OneDrive App Folder 안에 BluePage 폴더 생성 완료");
        return (driveId, _bluePageFolderId);
    }

    /// <summary>업로드된 파일이 실제로 쌓이는 OneDrive App Folder를 브라우저에서 열 수 있는 링크를 반환한다.</summary>
    public async Task<string> GetAppFolderWebUrlAsync(CancellationToken ct)
    {
        await GetBluePageFolderReferenceAsync(ct);
        return _bluePageFolderWebUrl ?? throw new InvalidOperationException("BluePage 폴더의 webUrl을 확인할 수 없습니다.");
    }

    public async Task<CloudFileMetadata> CreateAsync(string localFilePath, CancellationToken ct) =>
        ToCloudMetadata(await CreateInAppFolderAsync(localFilePath, ct));

    public async Task<CloudFileMetadata> UpdateAsync(string remoteFileId, string localFilePath, CancellationToken ct) =>
        ToCloudMetadata(await UpdateContentAsync(remoteFileId, localFilePath, ct));

    async Task<CloudFileMetadata> ICloudDriveService.GetMetadataAsync(string remoteFileId, CancellationToken ct) =>
        ToCloudMetadata(await GetMetadataAsync(remoteFileId, ct));

    public Task DownloadAsync(string remoteFileId, string destinationPath, CancellationToken ct) =>
        DownloadContentAsync(remoteFileId, destinationPath, ct);

    public Task<string> GetBluePageFolderWebUrlAsync(CancellationToken ct) => GetAppFolderWebUrlAsync(ct);

    private static CloudFileMetadata ToCloudMetadata(DriveItem item) => new(
        item.Id ?? throw new InvalidOperationException("OneDrive 파일 ID가 비어 있습니다."),
        item.WebUrl ?? string.Empty,
        item.LastModifiedDateTime ?? DateTimeOffset.UtcNow);
}
