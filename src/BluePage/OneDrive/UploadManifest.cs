using System.Text.Json;
using System.Text.Json.Serialization;

namespace Microsoft365OfficeWebLauncher.OneDrive;

public sealed class ManifestEntry
{
    [JsonPropertyName("driveItemId")]
    public string DriveItemId { get; set; } = string.Empty;

    [JsonPropertyName("webUrl")]
    public string WebUrl { get; set; } = string.Empty;

    [JsonPropertyName("lastKnownLocalWriteUtc")]
    public DateTimeOffset LastKnownLocalWriteUtc { get; set; }

    [JsonPropertyName("lastKnownRemoteModifiedUtc")]
    public DateTimeOffset LastKnownRemoteModifiedUtc { get; set; }
}

/// <summary>
/// 로컬 파일 경로 ↔ OneDrive driveItemId 매핑을 %LOCALAPPDATA%\Microsoft365OfficeWebLauncher\manifest.json에 보관한다.
/// 동일 로컬 파일을 재실행할 때 새 사본을 만들지 않고 기존 온라인 항목을 재사용/갱신하기 위한 근거 데이터.
/// </summary>
public sealed class UploadManifest
{
    private static readonly JsonSerializerOptions SerializerOptions = new() { WriteIndented = true };

    private readonly string _path;
    private readonly Dictionary<string, ManifestEntry> _entries;

    private UploadManifest(string path, Dictionary<string, ManifestEntry> entries)
    {
        _path = path;
        _entries = entries;
    }

    public static UploadManifest Load()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft365OfficeWebLauncher");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "manifest.json");

        Dictionary<string, ManifestEntry>? entries = null;
        if (File.Exists(path))
        {
            try
            {
                entries = JsonSerializer.Deserialize<Dictionary<string, ManifestEntry>>(File.ReadAllText(path));
            }
            catch (JsonException)
            {
                entries = null;
            }
        }

        return new UploadManifest(path, entries ?? new Dictionary<string, ManifestEntry>(StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>
    /// 디스크의 manifest.json을 다시 읽어 메모리 내용을 최신화한다.
    /// 더블클릭으로 문서를 열 때는 매번 별도의 헤드리스 프로세스가 실행되어 같은 파일에 기록하므로,
    /// 오래 떠 있는 GUI 프로세스(트레이 상주, 백그라운드 폴링)는 동기화를 시도하기 전에 반드시 이걸 호출해
    /// 다른 프로세스가 이미 반영해 둔 최신 상태를 놓치지 않아야 한다(안 그러면 이미 처리된 변경을 다시
    /// "충돌"로 오판하게 된다).
    /// </summary>
    public void Reload()
    {
        if (!File.Exists(_path))
        {
            _entries.Clear();
            return;
        }

        try
        {
            var entries = JsonSerializer.Deserialize<Dictionary<string, ManifestEntry>>(File.ReadAllText(_path));
            _entries.Clear();
            if (entries is not null)
            {
                foreach (var (key, value) in entries)
                {
                    _entries[key] = value;
                }
            }
        }
        catch (JsonException)
        {
            // 파일이 손상돼 있으면 메모리에 있던 내용을 그대로 유지한다(깨진 내용으로 덮어쓰지 않음).
        }
    }

    public ManifestEntry? Get(string localFilePath) =>
        _entries.TryGetValue(NormalizeKey(localFilePath), out var entry) ? entry : null;

    public IReadOnlyDictionary<string, ManifestEntry> AllEntries => _entries;

    public void Set(string localFilePath, ManifestEntry entry) => _entries[NormalizeKey(localFilePath)] = entry;

    public void Remove(string localFilePath) => _entries.Remove(NormalizeKey(localFilePath));

    public void Save() => File.WriteAllText(_path, JsonSerializer.Serialize(_entries, SerializerOptions));

    private static string NormalizeKey(string localFilePath) =>
        Path.GetFullPath(localFilePath).ToLowerInvariant();
}
