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

    public async Task<CloudFileMetadata> CreateAsync(string localFilePath, CancellationToken ct)
    {
        var drive = await GetDriveAsync(ct);
        var folderId = await GetBluePageFolderIdAsync(ct);
        var metadata = new GoogleFile
        {
            Name = Path.GetFileName(localFilePath),
            Parents = new[] { folderId }
        };
        await using var stream = File.OpenRead(localFilePath);
        var request = drive.Files.Create(metadata, stream, GetMimeType(localFilePath));
        request.Fields = "id,webViewLink,modifiedTime";
        var progress = await request.UploadAsync(ct);
        if (progress.Status != Google.Apis.Upload.UploadStatus.Completed)
        {
            throw progress.Exception ?? new InvalidOperationException("Google Drive 업로드에 실패했습니다.");
        }
        _logger.Info($"Google Drive BluePage 폴더에 새로 업로드: {metadata.Name}");
        return ToMetadata(request.ResponseBody);
    }

    public async Task<CloudFileMetadata> UpdateAsync(string remoteFileId, string localFilePath, CancellationToken ct)
    {
        var drive = await GetDriveAsync(ct);
        await using var stream = File.OpenRead(localFilePath);
        var request = drive.Files.Update(new GoogleFile(), remoteFileId, stream, GetMimeType(localFilePath));
        request.Fields = "id,webViewLink,modifiedTime";
        var progress = await request.UploadAsync(ct);
        if (progress.Status != Google.Apis.Upload.UploadStatus.Completed)
        {
            throw progress.Exception ?? new InvalidOperationException("Google Drive 파일 갱신에 실패했습니다.");
        }
        _logger.Info($"기존 Google Drive 항목 갱신: {remoteFileId}");
        return ToMetadata(request.ResponseBody);
    }

    public async Task<CloudFileMetadata> GetMetadataAsync(string remoteFileId, CancellationToken ct)
    {
        var drive = await GetDriveAsync(ct);
        var request = drive.Files.Get(remoteFileId);
        request.Fields = "id,webViewLink,modifiedTime";
        return ToMetadata(await request.ExecuteAsync(ct));
    }

    public async Task DownloadAsync(string remoteFileId, string destinationPath, CancellationToken ct)
    {
        var drive = await GetDriveAsync(ct);
        var directory = Path.GetDirectoryName(destinationPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }
        await using var output = File.Create(destinationPath);
        var progress = await drive.Files.Get(remoteFileId).DownloadAsync(output, ct);
        if (progress.Status != Google.Apis.Download.DownloadStatus.Completed)
        {
            throw progress.Exception ?? new InvalidOperationException("Google Drive 다운로드에 실패했습니다.");
        }
    }

    public async Task<string> GetBluePageFolderWebUrlAsync(CancellationToken ct) =>
        $"https://drive.google.com/drive/folders/{await GetBluePageFolderIdAsync(ct)}";

    public Task<RemoteLockState> GetLockStateAsync(string remoteFileId, CancellationToken ct) =>
        Task.FromResult(RemoteLockState.Unknown);

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

    private async Task<string> GetBluePageFolderIdAsync(CancellationToken ct)
    {
        if (_folderId is not null)
        {
            return _folderId;
        }
        var drive = await GetDriveAsync(ct);
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
        ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xls" => "application/vnd.ms-excel",
        ".xlsm" => "application/vnd.ms-excel.sheet.macroEnabled.12",
        ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".ppt" => "application/vnd.ms-powerpoint",
        _ => "application/octet-stream"
    };
}
