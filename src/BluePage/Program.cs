using System.Threading;
using System.Windows.Forms;
using Microsoft365OfficeWebLauncher.Auth;
using Microsoft365OfficeWebLauncher.Config;
using Microsoft365OfficeWebLauncher.Core;
using Microsoft365OfficeWebLauncher.Logging;
using Microsoft365OfficeWebLauncher.OneDrive;
using Microsoft365OfficeWebLauncher.Registry;
using Microsoft365OfficeWebLauncher.UI;

namespace Microsoft365OfficeWebLauncher;

internal static class Program
{
    private sealed record AppServices(
        GraphAuthService AuthService,
        OneDriveUploadService UploadService,
        UploadManifest Manifest,
        LaunchOrchestrator Orchestrator,
        FileAssociationRegistrar Registrar,
        StartupRegistrar StartupRegistrar);

    [STAThread]
    private static int Main(string[] args)
    {
        ApplicationConfiguration.Initialize();

        var config = ConfigLoader.Load();
        var logger = new FileLogger(ParseLogLevel(config.Logging.Level), config.Logging.RetainDays);

        // 헤드리스(더블클릭 열기) 흐름에서도 진짜 충돌이면 ConflictResolutionDialog가 뜰 수 있으므로
        // GUI 여부와 무관하게 항상 먼저 테마를 초기화해 둔다.
        AppTheme.Initialize(ParseThemePreference(config.Theme));

        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            logger.Error("처리되지 않은 예외(AppDomain)", e.ExceptionObject as Exception ?? new Exception(e.ExceptionObject?.ToString()));
        };

        try
        {
            return RunAsync(args, config, logger).GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            logger.Error("실행 중 처리되지 않은 예외가 발생했습니다.", ex);
            MessageBox.Show(
                $"예기치 않은 오류가 발생했습니다.\n\n{ex.Message}\n\n자세한 내용은 로그를 확인하세요:\n{logger.CurrentLogFilePath}",
                AppBrand.Name,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }
    }

    private static async Task<int> RunAsync(string[] args, AppConfig config, FileLogger logger)
    {
        // 인수 없이 실행(시작 메뉴 등)하거나 --settings/--minimized로 실행하면 상태/설정 창을 띄운다.
        // --minimized는 Windows 자동 시작 등록 시 붙는 인수로, 창을 보여주지 않고 바로 트레이로 들어간다.
        // 문서 더블클릭(첫 인수가 파일 경로) 흐름은 아래에서 창 없이 그대로 처리된다.
        if (args.Length == 0
            || string.Equals(args[0], "--settings", StringComparison.OrdinalIgnoreCase)
            || string.Equals(args[0], "--minimized", StringComparison.OrdinalIgnoreCase))
        {
            // GUI(트레이 상주) 인스턴스는 하나만 허용한다. 이미 떠 있으면 새로 띄우는 대신
            // 기존 인스턴스에게 "창을 보여줘"라는 신호만 보내고 조용히 종료한다.
            using var instanceMutex = new Mutex(initiallyOwned: true, SingleInstance.MutexName, out var createdNew);
            if (!createdNew)
            {
                logger.Info("이미 실행 중인 인스턴스가 있어 새로 띄우지 않고 기존 창을 앞으로 가져옵니다.");
                try
                {
                    using var existingShowEvent = EventWaitHandle.OpenExisting(SingleInstance.ShowWindowEventName);
                    existingShowEvent.Set();
                }
                catch (WaitHandleCannotBeOpenedException)
                {
                    // 기존 인스턴스가 아직 이벤트를 준비하기 전이거나 막 종료하는 중인 아주 드문 경우 — 무시
                }
                return 0;
            }

            var startMinimized = args.Length > 0 && string.Equals(args[0], "--minimized", StringComparison.OrdinalIgnoreCase);
            var services = BuildServices(config, logger);
            var exePath = Environment.ProcessPath ?? throw new InvalidOperationException("실행 파일 경로를 확인할 수 없습니다.");
            Application.Run(new LauncherForm(
                config, logger, services.AuthService, services.UploadService, services.Manifest,
                services.Orchestrator, services.Registrar, services.StartupRegistrar, exePath, startMinimized));
            return 0;
        }

        var command = args[0];

        if (string.Equals(command, "--register", StringComparison.OrdinalIgnoreCase))
        {
            var exePath = Environment.ProcessPath ?? throw new InvalidOperationException("실행 파일 경로를 확인할 수 없습니다.");
            new FileAssociationRegistrar(logger).RegisterAll(config, exePath);
            MessageBox.Show(
                "파일 연결 등록이 완료되었습니다.\n" +
                $"탐색기에서 문서를 우클릭 → 연결 프로그램 → 다른 앱 선택에서 '{AppBrand.Name}'를 선택해\n" +
                "\"항상 이 앱을 사용\"으로 지정하면 더블클릭 시 자동으로 열립니다.\n" +
                "(Windows 정책상 앱이 스스로 기본값을 강제 지정할 수 없어 이 1회 확인이 필요합니다.)",
                "등록 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return 0;
        }

        if (string.Equals(command, "--unregister", StringComparison.OrdinalIgnoreCase))
        {
            new FileAssociationRegistrar(logger).UnregisterAll(config);
            MessageBox.Show("파일 연결 등록을 해제했습니다.", "해제 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return 0;
        }

        var orchestrator = BuildServices(config, logger).Orchestrator;

        if (string.Equals(command, "--sync-all", StringComparison.OrdinalIgnoreCase))
        {
            return await orchestrator.SyncAllAsync(CancellationToken.None);
        }

        if (string.Equals(command, "--sync", StringComparison.OrdinalIgnoreCase))
        {
            if (args.Length < 2)
            {
                MessageBox.Show("--sync 명령에는 파일 경로가 필요합니다.", "인수 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 1;
            }

            return await orchestrator.SyncOneAsync(args[1], CancellationToken.None);
        }

        // 그 외에는 첫 인수를 파일 경로로 취급한다 (탐색기 더블클릭 시 전달되는 형태)
        return await orchestrator.OpenAsync(command, CancellationToken.None);
    }

    private static AppServices BuildServices(AppConfig config, FileLogger logger)
    {
        var catalog = new DocumentTypeCatalog(config);
        var authService = new GraphAuthService(config, logger);
        var uploadService = new OneDriveUploadService(authService, logger);
        var manifest = UploadManifest.Load();
        var syncCoordinator = new SyncCoordinator(uploadService, manifest, new GuiConflictResolver(), logger);
        var orchestrator = new LaunchOrchestrator(catalog, syncCoordinator, manifest, logger);
        var registrar = new FileAssociationRegistrar(logger);
        var startupRegistrar = new StartupRegistrar();
        return new AppServices(authService, uploadService, manifest, orchestrator, registrar, startupRegistrar);
    }

    private static LogLevel ParseLogLevel(string level) => level.ToLowerInvariant() switch
    {
        "debug" => LogLevel.Debug,
        "warn" or "warning" => LogLevel.Warn,
        "error" => LogLevel.Error,
        _ => LogLevel.Info
    };

    private static ThemePreference ParseThemePreference(string theme) => theme.ToLowerInvariant() switch
    {
        "light" => ThemePreference.Light,
        "dark" => ThemePreference.Dark,
        _ => ThemePreference.System
    };
}
