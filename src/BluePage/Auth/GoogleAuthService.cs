using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using Microsoft365OfficeWebLauncher.Config;
using Microsoft365OfficeWebLauncher.Logging;

namespace Microsoft365OfficeWebLauncher.Auth;

public sealed class GoogleAuthService
{
    private readonly AppConfig _config;
    private readonly FileLogger _logger;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private UserCredential? _credential;
    private IDataStore? _dataStore;
    private ProtectedFileDataStore? _fileStore;
    private DateTime? _reauthRequiredAtUtc;

    public GoogleAuthService(AppConfig config, FileLogger logger)
    {
        _config = config;
        _logger = logger;
    }

    public bool IsConfigured =>
        !string.IsNullOrWhiteSpace(_config.GoogleAuth.ClientId) &&
        !string.IsNullOrWhiteSpace(_config.GoogleAuth.ClientSecret);

    /// <summary>
    /// 브라우저 동의 창을 띄워도 되는 흐름인지. 문서 열기/CLI처럼 사용자가 직접 시작한 프로세스에서만 true로 두고,
    /// 트레이 상주 인스턴스의 백그라운드 동기화에서는 false로 둬서 갑자기 로그인 창이 뜨지 않게 한다.
    /// </summary>
    public bool InteractiveAuthAllowed { get; set; }

    /// <summary>저장된 토큰이 만료/취소되어 사용자가 직접 다시 로그인해야 하는 상태인지.</summary>
    public bool ReauthenticationRequired => _reauthRequiredAtUtc is not null;

    public async Task<UserCredential> AcquireCredentialAsync(CancellationToken ct = default)
    {
        if (_credential is not null)
        {
            return _credential;
        }

        await _gate.WaitAsync(ct);
        try
        {
            return _credential ??= await AuthorizeAsync(ct);
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <summary>
    /// API 호출이 invalid_grant로 실패했을 때 호출한다.
    /// 1) 다른 프로세스가 이미 새로 로그인해 둔 토큰이 있으면 그것을 조용히 집어오고,
    /// 2) 대화형 로그인이 허용된 흐름이면 저장된 죽은 토큰을 지우고 동의 창을 띄운다.
    /// 둘 다 불가능하면 false를 돌려주고 다시 로그인이 필요한 상태로 표시한다(백그라운드 무한 재시도 방지).
    /// </summary>
    public async Task<bool> TryRecoverAsync(CancellationToken ct = default)
    {
        await _gate.WaitAsync(ct);
        try
        {
            var storedTokenWriteUtc = LatestTokenWriteUtc();
            _credential = null;

            // 다른 프로세스(문서 열기 등)가 방금 다시 로그인했다면 디스크의 토큰이 실패 시점보다 새것이다.
            if (_reauthRequiredAtUtc is { } failedAt && storedTokenWriteUtc > failedAt)
            {
                _logger.Info("다른 프로세스가 갱신한 Google 토큰을 발견해 다시 사용합니다.");
                _credential = await AuthorizeAsync(ct);
                _reauthRequiredAtUtc = null;
                return true;
            }

            if (!InteractiveAuthAllowed)
            {
                if (_reauthRequiredAtUtc is null)
                {
                    _reauthRequiredAtUtc = DateTime.UtcNow;
                    _logger.Warn("Google 리프레시 토큰이 만료/취소되었습니다. 다시 로그인하기 전까지 Google 동기화를 중단합니다. " +
                                 "(OAuth 동의 화면이 테스트 상태이면 리프레시 토큰은 7일마다 만료됩니다.)");
                }
                return false;
            }

            _logger.Info("Google 리프레시 토큰이 만료/취소되어 다시 로그인합니다.");
            ClearStoredToken();
            _credential = await AuthorizeAsync(ct);
            _reauthRequiredAtUtc = null;
            return true;
        }
        catch (Exception ex)
        {
            _reauthRequiredAtUtc ??= DateTime.UtcNow;
            _logger.Error("Google 재로그인 실패", ex);
            return false;
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <summary>
    /// 사용자가 직접 로그인 버튼을 눌렀을 때처럼, 저장된 토큰을 버리고 동의 창을 강제로 띄운다.
    /// </summary>
    public async Task<UserCredential> ForceInteractiveSignInAsync(CancellationToken ct = default)
    {
        await _gate.WaitAsync(ct);
        try
        {
            _credential = null;
            ClearStoredToken();
            _credential = await AuthorizeAsync(ct);
            _reauthRequiredAtUtc = null;
            return _credential;
        }
        finally
        {
            _gate.Release();
        }
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
        ClearStoredToken();
        if (_dataStore is MemoryDataStore memoryStore)
        {
            await memoryStore.ClearAsync();
        }
        ClearPersistedCache();
        _reauthRequiredAtUtc = null;
        _logger.Info("Google 로그아웃 완료");
    }

    public static void ClearPersistedCache()
    {
        if (Directory.Exists(TokenCacheDirectory))
        {
            Directory.Delete(TokenCacheDirectory, recursive: true);
        }
    }

    private async Task<UserCredential> AuthorizeAsync(CancellationToken ct)
    {
        if (!IsConfigured)
        {
            throw new InvalidOperationException(
                "Google 로그인을 사용하려면 config.json의 googleAuth.clientId와 clientSecret을 설정해야 합니다.");
        }

        if (_dataStore is null)
        {
            if (_config.SharedPcMode)
            {
                _dataStore = new MemoryDataStore();
            }
            else
            {
                _fileStore = new ProtectedFileDataStore(TokenCacheDirectory, _logger);
                _dataStore = _fileStore;
            }
        }

        var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
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
        return credential;
    }

    /// <summary>죽은 토큰은 우리가 의도한 순간에만 지운다(평소에는 ProtectedFileDataStore가 삭제를 막는다).</summary>
    private void ClearStoredToken()
    {
        if (_fileStore is null)
        {
            return;
        }

        using (_fileStore.AllowDeleteScope())
        {
            _fileStore.ClearAsync().GetAwaiter().GetResult();
        }
    }

    private static DateTime? LatestTokenWriteUtc()
    {
        if (!Directory.Exists(TokenCacheDirectory))
        {
            return null;
        }

        var writes = Directory.EnumerateFiles(TokenCacheDirectory)
            .Select(File.GetLastWriteTimeUtc)
            .ToList();
        return writes.Count == 0 ? null : writes.Max();
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
