using System.Text.Json;

namespace Microsoft365OfficeWebLauncher.Config;

/// <summary>
/// 번들된 appsettings.default.json을 최초 1회 %LOCALAPPDATA%\Microsoft365OfficeWebLauncher\config.json로 복사해
/// 관리자 권한 없이 사용자별로 편집 가능하게 하고, 이후 실행부터는 사용자 파일을 사용한다.
/// 향후 새 기본 문서 형식이 추가되면(신규 배포판의 default) 사용자 파일에 없는 확장자만 병합한다.
/// </summary>
public static class ConfigLoader
{
    private static readonly JsonSerializerOptions SerializerOptions = new() { WriteIndented = true };

    public static string UserConfigDirectory =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft365OfficeWebLauncher");

    public static string UserConfigPath => Path.Combine(UserConfigDirectory, "config.json");

    public static AppConfig Load()
    {
        var defaultConfig = LoadDefaultConfig();

        Directory.CreateDirectory(UserConfigDirectory);
        if (!File.Exists(UserConfigPath))
        {
            File.WriteAllText(UserConfigPath, JsonSerializer.Serialize(defaultConfig, SerializerOptions));
        }

        AppConfig userConfig;
        try
        {
            var json = File.ReadAllText(UserConfigPath);
            userConfig = JsonSerializer.Deserialize<AppConfig>(json, SerializerOptions) ?? defaultConfig;
        }
        catch (JsonException)
        {
            // 손상된 사용자 설정 파일은 기본값으로 대체(원본은 그대로 두어 사용자가 직접 확인 가능)
            userConfig = defaultConfig;
        }

        MergeMissingDocumentTypes(userConfig, defaultConfig);
        return userConfig;
    }

    /// <summary>사용자 설정 파일에 저장한다(GUI의 [저장] 버튼 등에서 사용).</summary>
    public static void Save(AppConfig config)
    {
        Directory.CreateDirectory(UserConfigDirectory);
        File.WriteAllText(UserConfigPath, JsonSerializer.Serialize(config, SerializerOptions));
    }

    private static AppConfig LoadDefaultConfig()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "appsettings.default.json");
        if (File.Exists(path))
        {
            var json = File.ReadAllText(path);
            var cfg = JsonSerializer.Deserialize<AppConfig>(json, SerializerOptions);
            if (cfg is not null)
            {
                return cfg;
            }
        }

        return new AppConfig();
    }

    private static void MergeMissingDocumentTypes(AppConfig target, AppConfig defaults)
    {
        var covered = new HashSet<string>(
            target.DocumentTypes.SelectMany(d => d.Extensions).Select(Normalize),
            StringComparer.OrdinalIgnoreCase);

        foreach (var defaultType in defaults.DocumentTypes)
        {
            var missing = defaultType.Extensions.Where(e => !covered.Contains(Normalize(e))).ToList();
            if (missing.Count == 0)
            {
                continue;
            }

            target.DocumentTypes.Add(new DocumentTypeConfig
            {
                Extensions = missing,
                OfficeApp = defaultType.OfficeApp
            });

            foreach (var ext in missing)
            {
                covered.Add(Normalize(ext));
            }
        }
    }

    private static string Normalize(string ext) => ext.StartsWith('.') ? ext : "." + ext;
}
