namespace Microsoft365OfficeWebLauncher.OneDrive;

/// <summary>로컬/온라인이 모두 변경된 충돌 상황에서 사용자가 선택하는 처리 방식.</summary>
public enum ConflictResolutionChoice
{
    /// <summary>온라인 사본을 별도 파일로 저장하고 로컬 원본은 그대로 둔다(기본값, 가장 안전).</summary>
    CreateCopy,

    /// <summary>수정 시각이 더 최근인 쪽으로 다른 쪽을 덮어쓴다.</summary>
    KeepNewer,

    /// <summary>로컬(오프라인) 파일 내용으로 온라인 사본을 덮어쓴다.</summary>
    KeepLocal,

    /// <summary>온라인 사본 내용으로 로컬 파일을 덮어쓴다.</summary>
    KeepRemote
}

/// <summary>충돌 해결 다이얼로그에 전달할 정보.</summary>
public sealed record ConflictInfo(string LocalFilePath, DateTimeOffset LocalModifiedUtc, DateTimeOffset RemoteModifiedUtc);

/// <summary>로컬/온라인 동시 변경(충돌) 시 어떻게 처리할지 사용자에게 물어보는 역할.</summary>
public interface IConflictResolver
{
    ConflictResolutionChoice Resolve(ConflictInfo info);
}
