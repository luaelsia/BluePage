using Google.Apis.Auth.OAuth2.Responses;

namespace Microsoft365OfficeWebLauncher.Auth;

/// <summary>
/// 저장된 Google 리프레시 토큰이 더 이상 유효하지 않아(만료/취소) 사용자가 직접 다시 로그인해야 하는 상태.
/// 대화형 로그인이 허용되지 않는 흐름(트레이 상주 인스턴스의 백그라운드 동기화 등)에서 던진다.
/// </summary>
public sealed class GoogleReauthRequiredException : Exception
{
    public GoogleReauthRequiredException()
        : base("Google 로그인이 만료되었습니다. BluePage 창에서 Google 계정에 다시 로그인해 주세요.")
    {
    }
}

public static class GoogleAuthErrorHelper
{
    /// <summary>
    /// 리프레시 토큰이 죽었을 때 Google이 돌려주는 invalid_grant 응답인지 확인한다.
    /// 토큰 갱신은 백그라운드 스레드에서 일어나 AggregateException/InnerException으로 감싸여 오는 경우가 많아
    /// 예외 체인 전체를 훑는다.
    /// </summary>
    public static bool IsInvalidGrant(Exception? ex)
    {
        while (ex is not null)
        {
            if (ex is TokenResponseException tokenEx &&
                string.Equals(tokenEx.Error?.Error, "invalid_grant", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (ex is AggregateException aggregate)
            {
                foreach (var inner in aggregate.Flatten().InnerExceptions)
                {
                    if (IsInvalidGrant(inner))
                    {
                        return true;
                    }
                }
                return false;
            }

            ex = ex.InnerException;
        }

        return false;
    }
}
