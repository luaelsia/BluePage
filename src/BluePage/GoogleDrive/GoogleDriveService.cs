using Google.Apis.Drive.v3;
using Google.Apis.Services;
using GoogleFile = Google.Apis.Drive.v3.Data.File;
using Microsoft365OfficeWebLauncher.Auth;
using Microsoft365OfficeWebLauncher.Cloud;
using Microsoft365OfficeWebLauncher.Logging;
using Microsoft365OfficeWebLauncher.OneDrive;

namespace Microsoft365OfficeWebLauncher.GoogleDrive;

public sealed class GoogleDriveService : ICloudDriveService
{
    private const string FolderMimeType = "application/vnd.google-apps.folder";
    private readonly GoogleAuthService _authService;
    private readonly FileLogger _logger;
    private DriveService? _drive;
    private string? _folderId;

    public GoogleDriveService(GoogleAuthService authService, FileLogger logger)
    {
        _authService = authService;
        _logger = logger;
    }

    public CloudProvider Provider => CloudProvider.Google;
    public string DisplayName => "Google Workspace";

    public Task<CloudFileMetadata> CreateAsync(string localFilePath, string remoteFileName, CancellationToken ct) =>
        ExecuteAsync(async (drive, token) =>
        {
            var folderId = await GetBluePageFolderIdAsync(drive, token);
            var metadata = new GoogleFile
            {
                Name = remoteFileName,
                Parents = new[] { folderId }
            };
            await using var stream = File.OpenRead(localFilePath);
            var request = drive.Files.Create(metadata, stream, GetMimeType(localFilePath));
            request.Fields = "id,webViewLink,modifiedTime";
            var progress = await request.UploadAsync(token);
            if (progress.Status != Google.Apis.Upload.UploadStatus.Completed)
            {
                throw progress.Exception ?? new InvalidOperationException("Google Drive 업로드에 실패했습니다.");
            }
            _logger.Info($"Google Drive BluePage 폴더에 새로 업로드: {metadata.Name}");
            return ToMetadata(request.ResponseBody);
        }, ct);

    public Task<CloudFileMetadata> UpdateAsync(string remoteFileId, string localFilePath, CancellationToken ct) =>
        ExecuteAsync(async (drive, token) =>
        {
            await using var stream = File.OpenRead(localFilePath);
            var request = drive.Files.Update(new GoogleFile(), remoteFileId, stream, GetMimeType(localFilePath));
            request.Fields = "id,webViewLink,modifiedTime";
            var progress = await request.UploadAsync(token);
            if (progress.Status != Google.Apis.Upload.UploadStatus.Completed)
            {
                throw progress.Exception ?? new InvalidOperationException("Google Drive 파일 갱신에 실패했습니다.");
            }
            _logger.Info($"기존 Google Drive 항목 갱신: {remoteFileId}");
            return ToMetadata(request.ResponseBody);
        }, ct);

    public Task<CloudFileMetadata> GetMetadataAsync(string remoteFileId, CancellationToken ct) =>
        ExecuteAsync(async (drive, token) =>
        {
            var request = drive.Files.Get(remoteFileId);
            request.Fields = "id,webViewLink,modifiedTime";
            return ToMetadata(await request.ExecuteAsync(token));
        }, ct);

    public Task DownloadAsync(string remoteFileId, string destinationPath, CancellationToken ct) =>
        ExecuteAsync(async (drive, token) =>
        {
            var directory = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }
            await using var output = File.Create(destinationPath);
            var progress = await drive.Files.Get(remoteFileId).DownloadAsync(output, token);
            if (progress.Status != Google.Apis.Download.DownloadStatus.Completed)
            {
                throw progress.Exception ?? new InvalidOperationException("Google Drive 다운로드에 실패했습니다.");
            }
            return true;
        }, ct);

    public Task<string> GetBluePageFolderWebUrlAsync(CancellationToken ct) =>
        ExecuteAsync(async (drive, token) =>
            $"https://drive.google.com/drive/folders/{await GetBluePageFolderIdAsync(drive, token)}", ct);

    public Task<RemoteLockState> GetLockStateAsync(string remoteFileId, CancellationToken ct) =>
        Task.FromResult(RemoteLockState.Unknown);

    /// <summary>
    /// Drive 호출을 실행하되, 저장된 리프레시 토큰이 죽어(invalid_grant) 실패하면 한 번만 복구를 시도한다.
    /// 복구는 GoogleAuthService가 판단한다 — 다른 프로세스가 새로 로그인해 둔 토큰을 집어오거나,
    /// 대화형 로그인이 허용된 흐름이면 동의 창을 띄운다. 복구가 불가능하면 GoogleReauthRequiredException을 던져
    /// 백그라운드 동기화가 같은 오류로 무한히 재시도하지 않게 한다.
    /// </summary>
    private async Task<T> ExecuteAsync<T>(Func<DriveService, CancellationToken, Task<T>> action, CancellationToken ct)
    {
        try
        {
            return await action(await GetDriveAsync(ct), ct);
        }
        catch (Exception ex) when (GoogleAuthErrorHelper.IsInvalidGrant(ex))
        {
            _drive = null;
            _folderId = null;

            if (!await _authService.TryRecoverAsync(ct))
            {
                throw new GoogleReauthRequiredException();
            }

            return await action(await GetDriveAsync(ct), ct);
        }
    }

    private async Task<DriveService> GetDriveAsync(CancellationToken ct)
    {
        if (_drive is not null)
        {
            return _drive;
        }
        var credential = await _authService.AcquireCredentialAsync(ct);
        _drive = new DriveService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = "BluePage"
        });
        return _drive;
    }

    private async Task<string> GetBluePageFolderIdAsync(DriveService drive, CancellationToken ct)
    {
        if (_folderId is not null)
        {
            return _folderId;
        }
        var list = drive.Files.List();
        list.Q = $"name = 'BluePage' and mimeType = '{FolderMimeType}' and trashed = false";
        list.Spaces = "drive";
        list.Fields = "files(id)";
        list.PageSize = 1;
        var existing = (await list.ExecuteAsync(ct)).Files?.FirstOrDefault();
        if (existing?.Id is not null)
        {
            _folderId = existing.Id;
            return _folderId;
        }
        var create = drive.Files.Create(new GoogleFile { Name = "BluePage", MimeType = FolderMimeType });
        create.Fields = "id";
        _folderId = (await create.ExecuteAsync(ct)).Id
            ?? throw new InvalidOperationException("Google Drive BluePage 폴더를 만들지 못했습니다.");
        _logger.Info("Google Drive에 BluePage 폴더 생성 완료");
        return _folderId;
    }

    private static CloudFileMetadata ToMetadata(GoogleFile file) => new(
        file.Id ?? throw new InvalidOperationException("Google Drive 파일 ID가 비어 있습니다."),
        file.WebViewLink ?? $"https://drive.google.com/open?id={file.Id}",
        file.ModifiedTimeDateTimeOffset ?? DateTimeOffset.UtcNow);

    private static string GetMimeType(string path) => Path.GetExtension(path).ToLowerInvariant() switch
    {
        ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".doc" => "application/msword",
        ".odt" => "application/vnd.oasis.opendocument.text",
        ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xls" => "application/vnd.ms-excel",
        ".xlsm" => "application/vnd.ms-excel.sheet.macroEnabled.12",
        ".csv" => "text/csv",
        ".ods" => "application/vnd.oasis.opendocument.spreadsheet",
        ".tsv" => "text/tab-separated-values",
        ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".ppt" => "application/vnd.ms-powerpoint",
        ".odp" => "application/vnd.oasis.opendocument.presentation",
        _ => "application/octet-stream"
    };
}
