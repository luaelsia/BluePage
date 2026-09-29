using Google.Apis.Util.Store;
using Microsoft365OfficeWebLauncher.Logging;

namespace Microsoft365OfficeWebLauncher.Auth;

/// <summary>
/// FileDataStore를 감싸 "라이브러리가 마음대로 토큰 파일을 지우는 것"만 막는 저장소.
///
/// 문서 더블클릭 열기는 항상 별도 프로세스라 디스크에 저장된 토큰에만 의존하는데,
/// 트레이 상주 인스턴스가 이미 죽은 리프레시 토큰으로 갱신을 시도하면 Google 클라이언트 라이브러리가
/// 실패 처리로 저장소의 토큰을 삭제해 버린다. 그러면 방금 다시 로그인해 저장한 새 토큰까지 같이 날아가
/// 파일을 열 때마다 로그인 창이 뜨게 된다.
/// 삭제는 우리가 의도한 순간(로그아웃, 재로그인 직전)에만 AllowDeleteScope 안에서 수행한다.
/// </summary>
internal sealed class ProtectedFileDataStore : IDataStore
{
    private readonly FileDataStore _inner;
    private readonly FileLogger _logger;
    private bool _deleteAllowed;

    public ProtectedFileDataStore(string folderPath, FileLogger logger)
    {
        _inner = new FileDataStore(folderPath, fullPath: true);
        _logger = logger;
    }

    public IDisposable AllowDeleteScope() => new DeletePermission(this);

    public Task StoreAsync<T>(string key, T value) => _inner.StoreAsync(key, value);

    public Task<T?> GetAsync<T>(string key) => _inner.GetAsync<T>(key)!;

    public Task DeleteAsync<T>(string key)
    {
        if (!_deleteAllowed)
        {
            _logger.Debug($"저장된 Google 토큰 삭제 요청을 무시했습니다(키: {key}).");
            return Task.CompletedTask;
        }
        return _inner.DeleteAsync<T>(key);
    }

    public Task ClearAsync()
    {
        if (!_deleteAllowed)
        {
            _logger.Debug("저장된 Google 토큰 전체 삭제 요청을 무시했습니다.");
            return Task.CompletedTask;
        }
        return _inner.ClearAsync();
    }

    private sealed class DeletePermission : IDisposable
    {
        private readonly ProtectedFileDataStore _store;

        public DeletePermission(ProtectedFileDataStore store)
        {
            _store = store;
            _store._deleteAllowed = true;
        }

        public void Dispose() => _store._deleteAllowed = false;
    }
}
