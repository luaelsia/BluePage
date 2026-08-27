using System.Text;

namespace Microsoft365OfficeWebLauncher.Logging;

/// <summary>
/// %LOCALAPPDATA%\Microsoft365OfficeWebLauncher\logs 에 일자별 롤링 로그를 기록하는 단순 파일 로거.
/// 외부 로깅 라이브러리 없이(요구사항: 외부 라이브러리 최소화) 직접 구현한다.
/// </summary>
public sealed class FileLogger
{
    private readonly object _lock = new();
    private readonly string _logDirectory;
    private readonly LogLevel _minLevel;
    private readonly int _retainDays;

    public FileLogger(LogLevel minLevel, int retainDays)
    {
        _minLevel = minLevel;
        _retainDays = retainDays <= 0 ? 14 : retainDays;
        _logDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft365OfficeWebLauncher",
            "logs");

        Directory.CreateDirectory(_logDirectory);
        CleanupOldLogs();
    }

    public string CurrentLogFilePath =>
        Path.Combine(_logDirectory, $"launcher-{DateTime.Now:yyyyMMdd}.log");

    public void Debug(string message) => Write(LogLevel.Debug, message);
    public void Info(string message) => Write(LogLevel.Info, message);
    public void Warn(string message) => Write(LogLevel.Warn, message);
    public void Error(string message) => Write(LogLevel.Error, message);

    public void Error(string message, Exception ex) =>
        Write(LogLevel.Error, $"{message}{Environment.NewLine}{ex}");

    private void Write(LogLevel level, string message)
    {
        if (level < _minLevel)
        {
            return;
        }

        var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level,-5}] {message}{Environment.NewLine}";

        lock (_lock)
        {
            try
            {
                File.AppendAllText(CurrentLogFilePath, line, Encoding.UTF8);
            }
            catch
            {
                // 로그 기록 실패는 앱 동작을 막지 않는다.
            }
        }
    }

    private void CleanupOldLogs()
    {
        try
        {
            var cutoff = DateTime.Now.AddDays(-_retainDays);
            foreach (var file in Directory.EnumerateFiles(_logDirectory, "launcher-*.log"))
            {
                if (File.GetLastWriteTime(file) < cutoff)
                {
                    File.Delete(file);
                }
            }
        }
        catch
        {
            // 정리 실패는 무시 — 다음 실행에서 재시도됨
        }
    }
}
