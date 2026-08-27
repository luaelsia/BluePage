namespace Microsoft365OfficeWebLauncher;

/// <summary>
/// GUI(트레이 상주) 인스턴스가 중복 실행되지 않도록 Program.cs와 LauncherForm이 함께 쓰는 이름들.
/// 문서를 더블클릭해서 여는 헤드리스 실행은 이 대상이 아니며 항상 별도 프로세스로 동시에 실행될 수 있다.
/// </summary>
internal static class SingleInstance
{
    public const string MutexName = "BluePage_SingleInstance";
    public const string ShowWindowEventName = "BluePage_ShowWindow";
}
