using System.Drawing;
using System.Threading;
using System.Windows;
using System.Windows.Threading;

namespace AudioDeviceSwitcher;

public partial class App : Application
{
    private const string MutexName = "AudioDeviceSwitcher-SingleInstance-{B7A2F4E1-9D3C-4E8A-A6B5-2F7D8E1C4A3B}";
    private const string ShowSignalName = "AudioDeviceSwitcher-Show-{B7A2F4E1-9D3C-4E8A-A6B5-2F7D8E1C4A3B}";

    private Mutex? _singleInstanceMutex;
    private EventWaitHandle? _showSignal;
    private System.Windows.Forms.NotifyIcon _trayIcon = null!;
    private MainWindow? _mainWindow;
    private DispatcherTimer? _deviceWatchTimer;
    private DeviceChangeNotifier? _deviceNotifier;
    private DateTime _suppressDeviceBalloonUntil = DateTime.MinValue;
    private string? _knownPlaybackId;
    private string? _knownRecordingId;
    private string? _knownPlaybackCommId;
    private string? _knownRecordingCommId;
    private bool _knownBluetooth;
    private bool _firstPoll = true;
    private HashSet<string> _knownDriftedApps = new(StringComparer.OrdinalIgnoreCase);

    public IReadOnlyCollection<string> DriftedApps => _knownDriftedApps;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Crash handlers first so any init failure below is captured.
        CrashLogger.Initialize();

        // Initialize AUMID + registry early so toasts fire under the right identity.
        ToastService.Initialize();

        _singleInstanceMutex = new Mutex(true, MutexName, out bool isFirstInstance);
        if (!isFirstInstance)
        {
            // Another instance is running — signal it to show, then exit
            try
            {
                using var signal = EventWaitHandle.OpenExisting(ShowSignalName);
                signal.Set();
            }
            catch { }
            Shutdown();
            return;
        }

        _showSignal = new EventWaitHandle(false, EventResetMode.AutoReset, ShowSignalName);
        ThreadPool.RegisterWaitForSingleObject(_showSignal, (_, _) =>
        {
            Dispatcher.Invoke(ShowMainWindow);
        }, null, -1, false);

        ShutdownMode = ShutdownMode.OnExplicitShutdown;

        _trayIcon = new System.Windows.Forms.NotifyIcon
        {
            Icon = SystemIcons.Application,
            Text = "音频切换助手",
            Visible = true,
        };

        try
        {
            var exePath = Environment.ProcessPath;
            if (exePath != null)
            {
                var icon = Icon.ExtractAssociatedIcon(exePath);
                if (icon != null) _trayIcon.Icon = icon;
            }
        }
        catch { }

        var menu = new System.Windows.Forms.ContextMenuStrip();
        menu.Items.Add("打开主窗口", null, (_, _) => ShowMainWindow());
        menu.Items.Add("迷你窗口", null, (_, _) => ShowMiniFromTray());
        menu.Items.Add("设置…", null, (_, _) => ShowSettingsWindow());
        menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());

        var autoStartItem = new System.Windows.Forms.ToolStripMenuItem("开机自启")
        {
            CheckOnClick = true,
            Checked = AutoStartService.IsEnabled(),
        };
        autoStartItem.CheckedChanged += (_, _) =>
        {
            try
            {
                AutoStartService.SetEnabled(autoStartItem.Checked);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置开机自启失败：\n{ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                autoStartItem.Checked = AutoStartService.IsEnabled();
            }
        };
        menu.Items.Add(autoStartItem);
        menu.Opening += (_, _) => autoStartItem.Checked = AutoStartService.IsEnabled();

        menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());
        menu.Items.Add("退出", null, (_, _) => ExitApp());
        _trayIcon.ContextMenuStrip = menu;

        _trayIcon.DoubleClick += (_, _) => ShowMainWindow();

        // Event-driven (instant) via Windows audio COM notifications.
        _deviceNotifier = new DeviceChangeNotifier(CheckDeviceChanges);

        // Low-frequency safety net: covers any missed events and refreshes app-drift state
        // (which depends on external process lifecycle, not just device COM events).
        _deviceWatchTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(5) };
        _deviceWatchTimer.Tick += (_, _) => CheckDeviceChanges();
        _deviceWatchTimer.Start();

        var startup = SettingsService.Load();
        if (startup.StartMinimized)
        {
            _mainWindow = new MainWindow();
            new System.Windows.Interop.WindowInteropHelper(_mainWindow).EnsureHandle();
        }
        else
        {
            ShowMainWindow();
        }

        if (startup.MiniWindowVisible)
        {
            if (_mainWindow == null) ShowMainWindow();
            _mainWindow?.OpenMiniWindow();
        }
    }

    private void ShowSettingsWindow()
    {
        var owner = _mainWindow is { IsLoaded: true, IsVisible: true } ? _mainWindow : null;
        var dlg = new SettingsWindow { Owner = owner! };
        if (owner == null) dlg.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        if (dlg.ShowDialog() == true)
        {
            _mainWindow?.RefreshFromExternalChange();
        }
    }

    public void ShowBalloon(string title, string message, bool warning = false)
    {
        // Generic balloons use a dedicated tag so they don't clobber the more important per-category toasts.
        ToastService.Show("generic", title, message);
    }

    public void NotifyProfileApplied(DeviceProfile profile, ProfileApplyResult result)
    {
        if (!SettingsService.Load().NotifyProfileApplied) return;
        // Toasts with the same Tag auto-replace — no cooldown/debounce logic needed.
        ShowProfileAppliedBalloon(profile, result);
    }

    private void ShowProfileAppliedBalloon(DeviceProfile profile, ProfileApplyResult result)
    {
        var settings = SettingsService.Load();
        string Resolve(string? id, string? fallback)
        {
            var raw = fallback ?? "无";
            if (!string.IsNullOrEmpty(id)
                && settings.DeviceNicknames.TryGetValue(id, out var nn)
                && !string.IsNullOrEmpty(nn))
                return $"{nn} ({raw})";
            return raw;
        }

        var body = $"播放: {Resolve(profile.PlaybackDeviceId, profile.PlaybackDeviceName)}\n"
                 + $"录音: {Resolve(profile.RecordingDeviceId, profile.RecordingDeviceName)}";
        bool warning = result.SkippedAppNames.Count > 0;
        if (warning)
        {
            var names = string.Join(", ", result.SkippedAppNames.Distinct(StringComparer.OrdinalIgnoreCase));
            body += $"\n\n{result.SkippedAppNames.Count} 个应用未运行，已跳过: {names}";
        }
        ToastService.Show(ToastService.TagProfileSwitch, $"已切换到 {profile.Name}", body);
    }

    public void MarkOwnChange()
    {
        var playback = AudioDeviceService.GetPlaybackDevices().Find(d => d.IsDefault);
        var recording = AudioDeviceService.GetRecordingDevices().Find(d => d.IsDefault);
        _knownPlaybackId = playback?.Id;
        _knownRecordingId = recording?.Id;
        _knownPlaybackCommId = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Render).Id;
        _knownRecordingCommId = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Capture).Id;
        // Suppress the "device changed" balloon for a short window — COM events from our own
        // change may arrive after MarkOwnChange updates state, so this guards the race.
        _suppressDeviceBalloonUntil = DateTime.Now.AddMilliseconds(1500);
    }

    private void CheckDeviceChanges()
    {
        // Always refresh MainWindow so bluetooth warning and profile match are up-to-date
        _mainWindow?.RefreshFromExternalChange();

        var playback = AudioDeviceService.GetPlaybackDevices().Find(d => d.IsDefault);
        var recording = AudioDeviceService.GetRecordingDevices().Find(d => d.IsDefault);
        var playbackComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Render);
        var recordingComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Capture);
        var bluetooth = AudioDeviceService.HasBluetoothDevice();

        if (_firstPoll)
        {
            _knownPlaybackId = playback?.Id;
            _knownRecordingId = recording?.Id;
            _knownPlaybackCommId = playbackComm.Id;
            _knownRecordingCommId = recordingComm.Id;
            _knownBluetooth = bluetooth;
            _firstPoll = false;
            return;
        }

        var settings = SettingsService.Load();

        if (bluetooth != _knownBluetooth)
        {
            if (settings.NotifyBluetooth)
            {
                ToastService.Show(ToastService.TagBluetooth,
                    bluetooth ? "\u84DD\u7259\u8BBE\u5907\u5DF2\u8FDE\u63A5" : "\u84DD\u7259\u8BBE\u5907\u5DF2\u65AD\u5F00",
                    bluetooth ? "\u5DE5\u4F5C\u65F6\u8BF7\u6CE8\u610F\u65AD\u5F00\u84DD\u7259" : "\u5F53\u524D\u65E0\u84DD\u7259\u97F3\u9891\u8BBE\u5907");
            }
            _knownBluetooth = bluetooth;
        }

        var playbackChanged = playback?.Id != _knownPlaybackId;
        var recordingChanged = recording?.Id != _knownRecordingId;
        var playbackCommChanged = playbackComm.Id != _knownPlaybackCommId;
        var recordingCommChanged = recordingComm.Id != _knownRecordingCommId;

        if (playbackChanged || recordingChanged || playbackCommChanged || recordingCommChanged)
        {
            bool inSuppressionWindow = DateTime.Now < _suppressDeviceBalloonUntil;
            if (settings.NotifyDeviceChanged && !inSuppressionWindow)
            {
                var lines = new List<string>();
                if (playbackChanged) lines.Add($"\u64AD\u653E: {playback?.Name ?? "\u65E0"}");
                if (playbackCommChanged && playbackComm.Id != playback?.Id)
                    lines.Add($"\u64AD\u653E(\u901A\u4FE1): {playbackComm.Name ?? "\u65E0"}");
                if (recordingChanged) lines.Add($"\u5F55\u97F3: {recording?.Name ?? "\u65E0"}");
                if (recordingCommChanged && recordingComm.Id != recording?.Id)
                    lines.Add($"\u5F55\u97F3(\u901A\u4FE1): {recordingComm.Name ?? "\u65E0"}");

                ToastService.Show(ToastService.TagDeviceChange,
                    "\u97F3\u9891\u8BBE\u5907\u5DF2\u66F4\u6539",
                    string.Join("\n", lines));
            }

            _knownPlaybackId = playback?.Id;
            _knownRecordingId = recording?.Id;
            _knownPlaybackCommId = playbackComm.Id;
            _knownRecordingCommId = recordingComm.Id;
        }

        CheckAppOverrideDrift(playback?.Id, recording?.Id);
    }

    private void CheckAppOverrideDrift(string? currentPlaybackId, string? currentRecordingId)
    {
        var profiles = ProfileService.GetAll();
        var active = profiles.Find(p =>
            p.PlaybackDeviceId == currentPlaybackId && p.RecordingDeviceId == currentRecordingId);
        if (active == null || active.AppOverrides.Count == 0)
        {
            _knownDriftedApps.Clear();
            return;
        }

        // Single batched Process.GetProcesses() call covers all overrides — beats one
        // full system scan per override (Process.GetProcesses is allocation-heavy).
        var allExes = active.AppOverrides
            .Select(o => o.ExePath)
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToHashSet(StringComparer.OrdinalIgnoreCase)!;
        var runningByPath = ProfileApplyService.GetRunningProcessesByPath(allExes);

        var drifted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var ov in active.AppOverrides)
        {
            if (string.IsNullOrWhiteSpace(ov.ExePath)) continue;
            var appProfile = AppProfileService.Get(ov.AppProfileId);
            if (appProfile == null) continue;

            if (!runningByPath.TryGetValue(ov.ExePath, out var pids) || pids.Count == 0) continue; // not running — skip

            // Query across all matching PIDs; return first non-null hit (some PIDs return
            // E_INVALIDARG for this app's identity mapping, others succeed).
            string? actualOut = null, actualIn = null;
            foreach (var pid in pids)
            {
                if (actualOut == null)
                {
                    try { actualOut = AppAudioRoutingService.GetAppEndpoint(pid, NAudio.CoreAudioApi.DataFlow.Render); } catch { }
                }
                if (actualIn == null)
                {
                    try { actualIn = AppAudioRoutingService.GetAppEndpoint(pid, NAudio.CoreAudioApi.DataFlow.Capture); } catch { }
                }
                if (actualOut != null && actualIn != null) break;
            }

            bool outOk = string.Equals(actualOut ?? "", appProfile.OutputDeviceId ?? "", StringComparison.OrdinalIgnoreCase);
            bool inOk = string.Equals(actualIn ?? "", appProfile.InputDeviceId ?? "", StringComparison.OrdinalIgnoreCase);
            if (!outOk || !inOk)
                drifted.Add(System.IO.Path.GetFileName(ov.ExePath));
        }

        // Only notify on newly-detected drift (entries appearing since last check)
        var newlyDrifted = drifted.Except(_knownDriftedApps, StringComparer.OrdinalIgnoreCase).ToList();
        _knownDriftedApps = drifted;

        if (newlyDrifted.Count > 0 && SettingsService.Load().NotifyAppDrift)
        {
            ToastService.Show(ToastService.TagAppDrift,
                $"{newlyDrifted.Count} \u4E2A\u5E94\u7528\u5DF2\u504F\u79BB\u914D\u7F6E",
                string.Join(", ", newlyDrifted));
        }
    }

    private void ShowMainWindow()
    {
        if (_mainWindow is { IsLoaded: true })
        {
            _mainWindow.Show();
            _mainWindow.WindowState = WindowState.Normal;
            BringToFront(_mainWindow);
            return;
        }

        _mainWindow = new MainWindow();
        _mainWindow.Show();
    }

    private static void BringToFront(Window window)
    {
        // Topmost flicker to force foreground (Windows blocks cross-process Activate)
        window.Topmost = true;
        window.Topmost = false;
        window.Activate();
        window.Focus();
    }

    private void ShowMiniFromTray()
    {
        if (_mainWindow == null || !_mainWindow.IsLoaded)
            ShowMainWindow();

        _mainWindow!.OpenMiniWindow();
        _mainWindow.Hide();
    }

    public void ExitFromUI() => ExitApp();

    private void ExitApp()
    {
        _deviceWatchTimer?.Stop();
        _deviceNotifier?.Dispose();
        _trayIcon.Visible = false;
        _trayIcon.Dispose();
        _mainWindow?.ForceClose();
        _showSignal?.Dispose();
        _singleInstanceMutex?.ReleaseMutex();
        _singleInstanceMutex?.Dispose();
        Shutdown();
    }
}
