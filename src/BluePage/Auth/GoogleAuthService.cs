using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using Microsoft365OfficeWebLauncher.Config;
using Microsoft365OfficeWebLauncher.Logging;

namespace Microsoft365OfficeWebLauncher.Auth;

public sealed class GoogleAuthService
{
    private readonly AppConfig _config;
    private readonly FileLogger _logger;
    private UserCredential? _credential;
    private IDataStore? _dataStore;

    public GoogleAuthService(AppConfig config, FileLogger logger)
    {
        _config = config;
        _logger = logger;
    }

    public bool IsConfigured =>
        !string.IsNullOrWhiteSpace(_config.GoogleAuth.ClientId) &&
        !string.IsNullOrWhiteSpace(_config.GoogleAuth.ClientSecret);

    public async Task<UserCredential> AcquireCredentialAsync(CancellationToken ct = default)
    {
        if (_credential is not null)
        {
            return _credential;
        }

        if (!IsConfigured)
        {
            throw new InvalidOperationException(
                "Google 로그인을 사용하려면 config.json의 googleAuth.clientId와 clientSecret을 설정해야 합니다.");
        }

        _dataStore = _config.SharedPcMode
            ? new MemoryDataStore()
            : new FileDataStore(TokenCacheDirectory, fullPath: true);

        _credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
            new ClientSecrets
            {
                ClientId = _config.GoogleAuth.ClientId,
                ClientSecret = _config.GoogleAuth.ClientSecret
            },
            _config.GoogleAuth.Scopes,
            "BluePage",
            ct,
            _dataStore);

        _logger.Info("Google 대화형/캐시 로그인 성공");
        return _credential;
    }

    public Task<bool> HasCachedAccountAsync()
    {
        if (_credential is not null)
        {
            return Task.FromResult(true);
        }

        return Task.FromResult(!_config.SharedPcMode &&
            Directory.Exists(TokenCacheDirectory) &&
            Directory.EnumerateFiles(TokenCacheDirectory).Any());
    }

    public async Task SignOutAsync(CancellationToken ct = default)
    {
        if (_credential is not null)
        {
            try
            {
                await _credential.RevokeTokenAsync(ct);
            }
            catch (Exception ex)
            {
                _logger.Warn($"Google 토큰 폐기 요청 실패(로컬 로그인 정보는 삭제): {ex.Message}");
            }
        }

        _credential = null;
        if (_dataStore is not null)
        {
            await _dataStore.ClearAsync();
        }
        ClearPersistedCache();
        _logger.Info("Google 로그아웃 완료");
    }

    public static void ClearPersistedCache()
    {
        if (Directory.Exists(TokenCacheDirectory))
        {
            Directory.Delete(TokenCacheDirectory, recursive: true);
        }
    }

    private static string TokenCacheDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Microsoft365OfficeWebLauncher",
        "google-token-cache");

    private sealed class MemoryDataStore : IDataStore
    {
        private readonly Dictionary<string, string> _values = new();

        public Task StoreAsync<T>(string key, T value)
        {
            _values[key] = System.Text.Json.JsonSerializer.Serialize(value);
            return Task.CompletedTask;
        }

        public Task DeleteAsync<T>(string key)
        {
            _values.Remove(key);
            return Task.CompletedTask;
        }

        public Task<T?> GetAsync<T>(string key)
        {
            if (!_values.TryGetValue(key, out var json))
            {
                return Task.FromResult(default(T));
            }
            return Task.FromResult(System.Text.Json.JsonSerializer.Deserialize<T>(json));
        }

        public Task ClearAsync()
        {
            _values.Clear();
            return Task.CompletedTask;
        }
    }
}
