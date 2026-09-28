namespace Microsoft365OfficeWebLauncher.Cloud;

/// <summary>
/// 클라우드의 BluePage 폴더는 하위 구조가 없는 단일 폴더라서, 파일 이름만으로 업로드하면
/// 서로 다른 폴더에 있는 동명 문서가 같은 원격 항목을 가리키게 된다. OneDrive는 같은 이름 경로에
/// replace로 올리기 때문에 먼저 올려둔 문서의 원격 내용이 실제로 덮어써지고, 그 뒤에는 매니페스트의
/// 두 로컬 경로가 같은 driveItemId를 공유해 한쪽 로컬 파일까지 상대 문서 내용으로 동기화된다.
///
/// 구분자로 로컬 경로 해시를 쓰면 안 된다. 같은 계정을 여러 PC에서 쓸 때 서로 다른 문서가
/// 우연히 같은 경로(예: C:\Users\사용자\문서\보스1기획서.docx)에 있으면 해시까지 같아져서
/// 원래 문제가 그대로 재현된다. 그래서 최초 업로드 시점에 난수로 발급해 매니페스트에 보관하는
/// 문서 ID(ManifestEntry.DocumentId)를 쓴다. 경로나 PC와 무관하게 문서마다 새로 발급되므로 충돌하지 않는다.
/// </summary>
public static class RemoteFileNaming
{
    /// <summary>예: 보스1기획서 (1a2b3c4d).docx</summary>
    public static string Build(string localFilePath, string documentId)
    {
        var name = Path.GetFileNameWithoutExtension(localFilePath);
        var extension = Path.GetExtension(localFilePath);
        return $"{name} ({documentId}){extension}";
    }

    /// <summary>새 문서 ID를 발급한다. 파일 이름에 넣어야 하므로 짧고 안전한 문자만 쓴다.</summary>
    public static string NewDocumentId() => Guid.NewGuid().ToString("N")[..8];
}
