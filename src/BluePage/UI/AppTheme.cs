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
    Color Failure);

/// <summary>
/// 다크 모드 상태를 앱 전체에서 공유하는 곳. 기본값은 "시스템 설정 따라가기"이며,
/// Windows 개인 설정(라이트/다크)이 바뀌면 <see cref="Changed"/>를 통해 열려 있는 창들이 즉시 반영한다.
/// </summary>
public static class AppTheme
{
    private static readonly ThemePalette LightPalette = new(
        Background: Color.FromArgb(243, 243, 243),
        CardBackground: Color.FromArgb(255, 255, 255),
        Border: Color.FromArgb(218, 218, 218),
        TextPrimary: Color.FromArgb(26, 26, 26),
        TextSecondary: Color.FromArgb(96, 96, 96),
        ButtonBackground: Color.FromArgb(251, 251, 251),
        ButtonBorder: Color.FromArgb(192, 192, 192),
        ButtonHover: Color.FromArgb(238, 238, 238),
        Accent: Color.FromArgb(0, 95, 184),
        Link: Color.FromArgb(15, 118, 110),
        Success: Color.FromArgb(16, 124, 16),
        Failure: Color.FromArgb(196, 43, 28));

    private static readonly ThemePalette DarkPalette = new(
        Background: Color.FromArgb(32, 32, 32),
        CardBackground: Color.FromArgb(43, 43, 43),
        Border: Color.FromArgb(70, 70, 70),
        TextPrimary: Color.FromArgb(255, 255, 255),
        TextSecondary: Color.FromArgb(197, 197, 197),
        ButtonBackground: Color.FromArgb(51, 51, 51),
        ButtonBorder: Color.FromArgb(85, 85, 85),
        ButtonHover: Color.FromArgb(66, 66, 66),
        Accent: Color.FromArgb(76, 194, 255),
        Link: Color.FromArgb(45, 212, 191),
        Success: Color.FromArgb(108, 203, 95),
        Failure: Color.FromArgb(255, 153, 164));

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
