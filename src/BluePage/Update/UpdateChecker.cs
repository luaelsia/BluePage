using System.Diagnostics;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Microsoft365OfficeWebLauncher.Update;

/// <summary>GitHub Releases의 최신 릴리스 정보. 설치 파일이 첨부되지 않았으면 InstallerUrl은 null.</summary>
public sealed record UpdateInfo(
    Version Version,
    string Tag,
    string ReleasePageUrl,
    string ReleaseNotes,
    string? InstallerName,
    string? InstallerUrl,
    long InstallerSize,
    string? InstallerSha256)
{
    public string DisplayVersion => $"{Version.Major}.{Version.Minor}.{Math.Max(0, Version.Build)}";
}

/// <summary>
/// 공개 저장소의 releases/latest API로 새 버전을 확인하고, 설치 파일을 받아 무인 설치로 실행한다.
/// 인증 없이 호출하므로 시간당 60회 제한이 있지만 하루 한 번 확인에는 충분하다.
/// </summary>
public static class UpdateChecker
{
    public const string RepositoryOwner = "luaelsia";
    public const string RepositoryName = "BluePage";
    public static string ReleasesPageUrl => $"https://github.com/{RepositoryOwner}/{RepositoryName}/releases";

    private static readonly Regex InstallerNamePattern =
        new(@"^BluePage-Setup-v[\d.]+\.exe$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly HttpClient Http = CreateClient();

    public static Version CurrentVersion
    {
        get
        {
            var version = typeof(UpdateChecker).Assembly.GetName().Version ?? new Version(0, 0, 0);
            return new Version(version.Major, version.Minor, Math.Max(0, version.Build));
        }
    }

    /// <summary>최신 릴리스를 가져온다. 현재 버전보다 높지 않으면 null.</summary>
    public static async Task<UpdateInfo?> GetNewerReleaseAsync(CancellationToken ct)
    {
        var latest = await GetLatestReleaseAsync(ct);
        return latest is not null && latest.Version > CurrentVersion ? latest : null;
    }

    public static async Task<UpdateInfo?> GetLatestReleaseAsync(CancellationToken ct)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(10));

        var url = $"https://api.github.com/repos/{RepositoryOwner}/{RepositoryName}/releases/latest";
        using var response = await Http.GetAsync(url, timeout.Token);
        response.EnsureSuccessStatusCode();

        await using var stream = await response.Content.ReadAsStreamAsync(timeout.Token);
        using var document = await JsonDocument.ParseAsync(stream, cancellationToken: timeout.Token);
        var root = document.RootElement;

        var tag = root.GetProperty("tag_name").GetString() ?? string.Empty;
        if (!TryParseVersion(tag, out var version))
        {
            return null;
        }

        string? installerName = null, installerUrl = null, installerSha256 = null;
        long installerSize = 0;
        if (root.TryGetProperty("assets", out var assets))
        {
            foreach (var asset in assets.EnumerateArray())
            {
                var name = asset.GetProperty("name").GetString() ?? string.Empty;
                if (!InstallerNamePattern.IsMatch(name))
                {
                    continue;
                }
                installerName = name;
                installerUrl = asset.GetProperty("browser_download_url").GetString();
                installerSize = asset.TryGetProperty("size", out var size) ? size.GetInt64() : 0;
                // GitHub가 계산해 두는 "sha256:..." 값. 오래된 릴리스에는 없을 수 있다.
                if (asset.TryGetProperty("digest", out var digest) &&
                    digest.GetString() is { } digestText &&
                    digestText.StartsWith("sha256:", StringComparison.OrdinalIgnoreCase))
                {
                    installerSha256 = digestText["sha256:".Length..];
                }
                break;
            }
        }

        return new UpdateInfo(
            version,
            tag,
            root.TryGetProperty("html_url", out var htmlUrl) ? htmlUrl.GetString() ?? ReleasesPageUrl : ReleasesPageUrl,
            root.TryGetProperty("body", out var body) ? body.GetString() ?? string.Empty : string.Empty,
            installerName,
            installerUrl,
            installerSize,
            installerSha256);
    }

    /// <summary>설치 파일을 임시 폴더에 받고 크기와 SHA-256을 확인한다. 확인에 실패하면 파일을 지우고 예외를 던진다.</summary>
    public static async Task<string> DownloadInstallerAsync(UpdateInfo update, IProgress<double>? progress, CancellationToken ct)
    {
        if (update.InstallerUrl is null || update.InstallerName is null)
        {
            throw new InvalidOperationException("이 릴리스에는 설치 파일이 첨부되어 있지 않습니다.");
        }

        var directory = Path.Combine(Path.GetTempPath(), "BluePage-Update");
        if (Directory.Exists(directory))
        {
            // 이전에 받다가 만 파일이나 지난 버전 설치 파일은 정리한다.
            foreach (var old in Directory.EnumerateFiles(directory))
            {
                try { File.Delete(old); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }
        Directory.CreateDirectory(directory);
        var destination = Path.Combine(directory, update.InstallerName);

        try
        {
            using (var response = await Http.GetAsync(update.InstallerUrl, HttpCompletionOption.ResponseHeadersRead, ct))
            {
                response.EnsureSuccessStatusCode();
                var total = response.Content.Headers.ContentLength ?? update.InstallerSize;

                await using var source = await response.Content.ReadAsStreamAsync(ct);
                await using var target = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true);
                var buffer = new byte[81920];
                long received = 0;
                int read;
                while ((read = await source.ReadAsync(buffer, ct)) > 0)
                {
                    await target.WriteAsync(buffer.AsMemory(0, read), ct);
                    received += read;
                    if (total > 0)
                    {
                        progress?.Report((double)received / total);
                    }
                }
            }

            var length = new FileInfo(destination).Length;
            if (update.InstallerSize > 0 && length != update.InstallerSize)
            {
                throw new InvalidDataException($"받은 파일 크기가 다릅니다(예상 {update.InstallerSize}, 실제 {length}).");
            }

            if (update.InstallerSha256 is { } expected)
            {
                await using var file = File.OpenRead(destination);
                var actual = Convert.ToHexString(await SHA256.HashDataAsync(file, ct));
                if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("받은 파일의 SHA-256 값이 GitHub에 기록된 값과 다릅니다.");
                }
            }

            return destination;
        }
        catch
        {
            try { File.Delete(destination); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            throw;
        }
    }

    /// <summary>
    /// 무인 설치로 실행한다. 설치 파일이 실행 중인 Blue Page를 끝내고, 설치가 끝나면 다시 실행한다(BluePage.iss 참고).
    /// </summary>
    public static void LaunchSilentInstall(string installerPath)
    {
        Process.Start(new ProcessStartInfo(installerPath, "/SILENT /SUPPRESSMSGBOXES /NORESTART")
        {
            UseShellExecute = true
        });
    }

    private static bool TryParseVersion(string tag, out Version version)
    {
        var text = tag.TrimStart('v', 'V');
        if (Version.TryParse(text, out var parsed))
        {
            version = new Version(parsed.Major, parsed.Minor, Math.Max(0, parsed.Build));
            return true;
        }
        version = new Version(0, 0, 0);
        return false;
    }

    private static HttpClient CreateClient()
    {
        var client = new HttpClient { Timeout = TimeSpan.FromMinutes(10) };
        client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("BluePage", CurrentVersion.ToString()));
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        return client;
    }
}
