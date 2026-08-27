namespace Microsoft365OfficeWebLauncher.OneDrive;

/// <summary>동기화 검토 창에서 사용자가 파일별로 고를 수 있는 실행 동작.</summary>
public enum SyncAction
{
    /// <summary>이번 실행에서 이 파일은 아무것도 하지 않는다.</summary>
    Skip,
    /// <summary>온라인 사본을 로컬로 반영(로컬 덮어씀).</summary>
    PullRemoteToLocal,
    /// <summary>로컬 내용을 온라인으로 반영(온라인 덮어씀).</summary>
    PushLocalToRemote,
    /// <summary>온라인 사본을 별도 파일로 저장하고 로컬 원본은 보존(충돌 상황의 안전한 기본값).</summary>
    CreateConflictCopy
}

/// <summary>업로드/다운로드 없이 조회만 한 결과 — 동기화 검토 창에 표시할 정보.</summary>
public sealed record SyncDetection(
    string LocalFilePath,
    SyncState DetectedState,
    DateTimeOffset LocalWriteUtc,
    DateTimeOffset RemoteModifiedUtc,
    DateTimeOffset LastSyncedUtc);
