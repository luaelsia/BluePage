using Microsoft.Graph.Models.ODataErrors;

namespace Microsoft365OfficeWebLauncher.OneDrive;

/// <summary>
/// Office Web에서 편집 중인 파일에 로컬 내용을 덮어쓰려고 하면 OneDrive/SharePoint가 항상
/// HTTP 423(Locked)로 거부한다. 이 상황을 다른 오류와 구분해 사용자에게 이해 가능한 메시지를 주기 위한 헬퍼.
/// </summary>
public static class GraphErrorHelper
{
    public static bool IsResourceLocked(Exception ex) =>
        ex is ODataError odataError && odataError.ResponseStatusCode == 423;
}
