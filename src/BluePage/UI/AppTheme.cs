using System.Runtime.InteropServices;

namespace Microsoft365OfficeWebLauncher.UI;

public enum ThemePreference
{
    System,
    Light,
    Dark
}

/// <summary>화면 곳곳(설정 창, 대화상자, 토스트)이 공유하는 색상 팔레트.</summary>
public sealed record ThemePalette(
    Color Background,
    Color CardBackground,
    Color Border,
    Color TextPrimary,
    Color TextSecondary,
    Color ButtonBackground,
    Color ButtonBorder,
    Color ButtonHover,
    Color Accent,
    Color Link,
    Color Success,
    Color Failure,
    Color Sidebar,
    Color AccentSoft,
    Color AccentSoftText,
    Color OnAccent,
    Color SuccessSoft,
    Color SuccessSoftText,
    Color WarningSoft,
    Color WarningSoftText,
    Color DangerSoft,
    Color DangerSoftText,
    Color NeutralSoft);

/// <summary>
/// 다크 모드 상태를 앱 전체에서 공유하는 곳. 기본값은 "시스템 설정 따라가기"이며,
/// Windows 개인 설정(라이트/다크)이 바뀌면 <see cref="Changed"/>를 통해 열려 있는 창들이 즉시 반영한다.
/// </summary>
public static class AppTheme
{
    // 앱 아이콘의 색에서 가져온 팔레트: 청록(바다) = 주요 강조색, 코랄(노을) = 보조 강조색,
    // 크림(접힌 모서리) = 배경. 상태 색도 이 톤에 맞춰 채도를 낮췄다.
    private static readonly ThemePalette LightPalette = new(
        Background: Color.FromArgb(251, 247, 240),
        CardBackground: Color.FromArgb(255, 253, 249),
        Border: Color.FromArgb(234, 223, 205),
        TextPrimary: Color.FromArgb(47, 58, 58),
        TextSecondary: Color.FromArgb(110, 119, 117),
        ButtonBackground: Color.FromArgb(255, 253, 249),
        ButtonBorder: Color.FromArgb(220, 207, 186),
        ButtonHover: Color.FromArgb(244, 236, 223),
        Accent: Color.FromArgb(30, 130, 132),
        Link: Color.FromArgb(30, 122, 124),
        Success: Color.FromArgb(46, 125, 85),
        Failure: Color.FromArgb(192, 70, 61),
        Sidebar: Color.FromArgb(244, 236, 223),
        AccentSoft: Color.FromArgb(216, 238, 236),
        AccentSoftText: Color.FromArgb(22, 112, 111),
        OnAccent: Color.FromArgb(255, 255, 255),
        SuccessSoft: Color.FromArgb(221, 240, 230),
        SuccessSoftText: Color.FromArgb(31, 112, 72),
        WarningSoft: Color.FromArgb(252, 233, 214),
        WarningSoftText: Color.FromArgb(165, 88, 42),
        DangerSoft: Color.FromArgb(252, 227, 218),
        DangerSoftText: Color.FromArgb(184, 72, 63),
        NeutralSoft: Color.FromArgb(239, 231, 218));

    private static readonly ThemePalette DarkPalette = new(
        Background: Color.FromArgb(29, 38, 39),
        CardBackground: Color.FromArgb(37, 49, 50),
        Border: Color.FromArgb(52, 66, 63),
        TextPrimary: Color.FromArgb(237, 230, 216),
        TextSecondary: Color.FromArgb(155, 165, 162),
        ButtonBackground: Color.FromArgb(43, 56, 57),
        ButtonBorder: Color.FromArgb(62, 78, 77),
        ButtonHover: Color.FromArgb(51, 66, 66),
        Accent: Color.FromArgb(108, 195, 189),
        Link: Color.FromArgb(127, 207, 200),
        Success: Color.FromArgb(143, 211, 168),
        Failure: Color.FromArgb(244, 164, 147),
        Sidebar: Color.FromArgb(23, 32, 33),
        AccentSoft: Color.FromArgb(35, 65, 63),
        AccentSoftText: Color.FromArgb(143, 211, 204),
        OnAccent: Color.FromArgb(16, 48, 47),
        SuccessSoft: Color.FromArgb(36, 69, 54),
        SuccessSoftText: Color.FromArgb(159, 221, 182),
        WarningSoft: Color.FromArgb(74, 55, 39),
        WarningSoftText: Color.FromArgb(242, 190, 143),
        DangerSoft: Color.FromArgb(74, 46, 43),
        DangerSoftText: Color.FromArgb(244, 164, 147),
        NeutralSoft: Color.FromArgb(47, 59, 59));

    public static ThemePreference Preference { get; private set; } = ThemePreference.System;

    /// <summary>사용자가 테마를 바꾸거나(설정 화면), Windows 시스템 테마가 바뀌었을 때 발생.</summary>
    public static event Action? Changed;

    public static bool IsDark => Preference switch
    {
        ThemePreference.Dark => true,
        ThemePreference.Light => false,
        _ => IsSystemDarkMode()
    };

    public static ThemePalette Current => IsDark ? DarkPalette : LightPalette;

    public static void Initialize(ThemePreference preference)
    {
        Preference = preference;
    }

    public static void SetPreference(ThemePreference preference)
    {
        Preference = preference;
        Changed?.Invoke();
    }

    /// <summary>Windows 개인 설정이 바뀌었을 때(SystemEvents.UserPreferenceChanged) 호출. "시스템 설정" 모드일 때만 의미가 있다.</summary>
    public static void NotifySystemThemeChanged()
    {
        if (Preference == ThemePreference.System)
        {
            Changed?.Invoke();
        }
    }

    private static bool IsSystemDarkMode()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            return key?.GetValue("AppsUseLightTheme") is int value && value == 0;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>타이틀바까지 다크로 맞추는 Windows 10(1809+)/11 네이티브 API. 미지원 환경에서도 예외 없이 조용히 실패한다.</summary>
    public static void ApplyTitleBarTheme(IntPtr handle, bool dark)
    {
        const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
        var value = dark ? 1 : 0;
        DwmSetWindowAttribute(handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));

        // DWMWCP_ROUND: Windows 11의 기본 둥근 창 모서리를 명시적으로 요청한다.
        var cornerPreference = 2;
        DwmSetWindowAttribute(handle, DWMWA_WINDOW_CORNER_PREFERENCE, ref cornerPreference, sizeof(int));
    }

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int valueSize);
}
