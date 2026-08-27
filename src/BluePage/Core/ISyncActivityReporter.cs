namespace Microsoft365OfficeWebLauncher.Core;

/// <summary>
/// 동기화가 시작/종료될 때마다 알림(예: 우하단 토스트 창)에 보고하기 위한 인터페이스.
/// GUI 프로세스에서만 실제 구현을 주입하고, 헤드리스 실행(더블클릭 오픈)이나 CLI에서는
/// 아무 것도 하지 않는 <see cref="NullSyncActivityReporter"/>를 그대로 사용한다.
/// </summary>
public interface ISyncActivityReporter
{
    void ReportStarted(string localFilePath);

    void ReportCompleted(string localFilePath, bool success);
}

public sealed class NullSyncActivityReporter : ISyncActivityReporter
{
    public static readonly NullSyncActivityReporter Instance = new();

    private NullSyncActivityReporter()
    {
    }

    public void ReportStarted(string localFilePath)
    {
    }

    public void ReportCompleted(string localFilePath, bool success)
    {
    }
}
