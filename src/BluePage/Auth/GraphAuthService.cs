using System.Runtime.InteropServices;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;
using Microsoft.Identity.Client.Extensions.Msal;
using Microsoft365OfficeWebLauncher.Config;
using Microsoft365OfficeWebLauncher.Logging;

namespace Microsoft365OfficeWebLauncher.Auth;

/// <summary>
/// MSAL + WAM(Web Account Manager) 브로커로 Microsoft 365 계정 토큰을 획득한다.
/// 비밀번호는 앱에 저장하지 않는다: 개인 PC에서는 암호화된 토큰 캐시로 사일런트 SSO,
/// 공용 PC(sharedPcMode)에서는 캐시를 디스크에 남기지 않고 매번 대화형 로그인한다.
/// </summary>
public sealed class GraphAuthService
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    private readonly AppConfig _config;
    private readonly FileLogger _logger;
    private IPublicClientApplication? _pca;

    public GraphAuthService(AppConfig config, FileLogger logger)
    {
        _config = config;
        _logger = logger;
    }

    public async Task<AuthenticationResult> AcquireTokenAsync(CancellationToken cancellationToken = default)
    {
        var pca = await GetOrCreatePcaAsync();
        var scopes = _config.Auth.Scopes;

        var accounts = await pca.GetAccountsAsync();
        var account = accounts.FirstOrDefault();

        try
        {
            var silentBuilder = account is not null
                ? pca.AcquireTokenSilent(scopes, account)
                : pca.AcquireTokenSilent(scopes, PublicClientApplication.OperatingSystemAccount);

            var result = await silentBuilder.ExecuteAsync(cancellationToken);
            _logger.Info($"사일런트 로그인 성공 (account={result.Account?.Username})");
            return result;
        }
        catch (MsalUiRequiredException)
        {
            _logger.Info("사일런트 로그인 불가 — 대화형 로그인으로 전환합니다.");
        }
        catch (MsalException ex)
        {
            _logger.Warn($"사일런트 로그인 실패({ex.ErrorCode}) — 대화형 로그인으로 전환합니다.");
        }

        var interactiveResult = await pca.AcquireTokenInteractive(scopes)
            .WithParentActivityOrWindow(GetForegroundWindow())
            .WithPrompt(account is null ? Prompt.SelectAccount : Prompt.NoPrompt)
            .ExecuteAsync(cancellationToken);

        _logger.Info($"대화형 로그인 성공 (account={interactiveResult.Account?.Username})");
        return interactiveResult;
    }

    /// <summary>로그인 팝업 없이, 캐시된 계정이 있으면 표시용 사용자 이름만 반환한다(GUI 상태 표시용).</summary>
    public async Task<string?> TryGetSignedInAccountAsync(CancellationToken cancellationToken = default)
    {
        var pca = await GetOrCreatePcaAsync();
        var accounts = await pca.GetAccountsAsync();
        return accounts.FirstOrDefault()?.Username;
    }

    /// <summary>캐시된 모든 계정을 제거한다(GUI의 로그아웃, 공용 PC 모드 전환용).</summary>
    public async Task SignOutAsync(CancellationToken cancellationToken = default)
    {
        var pca = await GetOrCreatePcaAsync();
        foreach (var account in await pca.GetAccountsAsync())
        {
            await pca.RemoveAsync(account);
        }

        _logger.Info("로그아웃 완료 — 캐시된 계정을 모두 제거했습니다.");
    }

    private async Task<IPublicClientApplication> GetOrCreatePcaAsync()
    {
        if (_pca is not null)
        {
            return _pca;
        }

        var builder = PublicClientApplicationBuilder
            .Create(_config.Auth.ClientId)
            .WithAuthority(_config.Auth.Authority)
            .WithDefaultRedirectUri()
            .WithBroker(new BrokerOptions(BrokerOptions.OperatingSystems.Windows));

        var pca = builder.Build();

        if (!_config.SharedPcMode)
        {
            await RegisterPersistentCacheAsync(pca);
        }
        else
        {
            _logger.Info("공용 PC 모드: 토큰 캐시를 디스크에 저장하지 않습니다.");
        }

        _pca = pca;
        return pca;
    }

    private async Task RegisterPersistentCacheAsync(IPublicClientApplication pca)
    {
        try
        {
            Directory.CreateDirectory(TokenCacheDirectory);

            var storageProperties = new StorageCreationPropertiesBuilder("msal_token_cache.bin", TokenCacheDirectory)
                .WithLinuxUnprotectedFile() // Windows에서는 무시되고 DPAPI가 사용됨; 비-Windows 폴백 안전장치
                .Build();

            var cacheHelper = await MsalCacheHelper.CreateAsync(storageProperties);
            cacheHelper.RegisterCache(pca.UserTokenCache);
        }
        catch (Exception ex)
        {
            _logger.Warn($"토큰 캐시 영속화 초기화 실패 — 이번 실행은 캐시 없이 진행합니다: {ex.Message}");
        }
    }

    private static string TokenCacheDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Microsoft365OfficeWebLauncher",
        "tokencache");

    /// <summary>디스크에 저장된 토큰 캐시 파일을 삭제한다(공용 PC 모드로 전환할 때 사용).</summary>
    public static void ClearPersistedCache()
    {
        if (Directory.Exists(TokenCacheDirectory))
        {
            Directory.Delete(TokenCacheDirectory, recursive: true);
        }
    }
}
