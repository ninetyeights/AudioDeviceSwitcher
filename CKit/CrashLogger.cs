using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace AudioDeviceSwitcher;

// Writes unhandled exceptions to %AppData%\AudioDeviceSwitcher\crash.log so users can ship
// us their crash report. UI-thread exceptions are handled (not rethrown) so the app stays up.
public static class CrashLogger
{
    private static readonly string LogDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AudioDeviceSwitcher");
    private static readonly string LogPath = Path.Combine(LogDir, "crash.log");
    private static readonly object _lock = new();

    public static string LogFilePath => LogPath;

    public static void Initialize()
    {
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            Write("AppDomain.UnhandledException", e.ExceptionObject as Exception);
            // Fatal — can't prevent termination here, just log.
        };

        Application.Current.DispatcherUnhandledException += (_, e) =>
        {
            Write("DispatcherUnhandledException", e.Exception);
            try
            {
                MessageBox.Show(
                    $"\u7A0B\u5E8F\u9047\u5230\u4E00\u4E2A\u9519\u8BEF\uFF0C\u5DF2\u8BB0\u5F55\u5230\u65E5\u5FD7\u6587\u4EF6\u3002\n\n{e.Exception.Message}\n\n\u65E5\u5FD7\u4F4D\u7F6E\uFF1A{LogPath}",
                    "\u97F3\u9891\u5207\u6362\u52A9\u624B", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch { }
            e.Handled = true; // keep the app alive for recoverable UI errors
        };

        TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            Write("TaskScheduler.UnobservedTaskException", e.Exception);
            e.SetObserved();
        };
    }

    public static void Write(string source, Exception? ex)
    {
        if (ex == null) return;
        try
        {
            lock (_lock)
            {
                Directory.CreateDirectory(LogDir);
                var entry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{source}]\n{ex}\n\n";
                File.AppendAllText(LogPath, entry);
            }
        }
        catch { }
    }
}
