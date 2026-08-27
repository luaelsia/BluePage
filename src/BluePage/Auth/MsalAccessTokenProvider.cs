using Microsoft.Kiota.Abstractions.Authentication;

namespace Microsoft365OfficeWebLauncher.Auth;

/// <summary>Graph SDK(Kiota)가 요구하는 IAccessTokenProvider를 GraphAuthService(MSAL)로 구현한다.</summary>
public sealed class MsalAccessTokenProvider : IAccessTokenProvider
{
    private readonly GraphAuthService _authService;

    public MsalAccessTokenProvider(GraphAuthService authService)
    {
        _authService = authService;
    }

    public AllowedHostsValidator AllowedHostsValidator { get; } = new();

    public async Task<string> GetAuthorizationTokenAsync(
        Uri uri,
        Dictionary<string, object>? additionalAuthenticationContext = null,
        CancellationToken cancellationToken = default)
    {
        var result = await _authService.AcquireTokenAsync(cancellationToken);
        return result.AccessToken;
    }
}
