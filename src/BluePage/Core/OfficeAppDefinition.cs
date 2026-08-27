namespace Microsoft365OfficeWebLauncher.Core;

/// <summary>확장자 하나에 대응하는 Office Web 앱 정보(로그/표시용). 실제 열기 URL은 Graph의 webUrl을 그대로 사용한다.</summary>
public sealed record OfficeAppDefinition(string OfficeApp, IReadOnlyList<string> Extensions);
