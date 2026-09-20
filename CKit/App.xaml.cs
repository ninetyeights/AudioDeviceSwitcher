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
    private AppRouteMonitor? _appRoutes;
    public ProfileScheduleService? ProfileSchedules { get; private set; }
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
        menu.Items.Add("打开主窗口", CreateMenuIcon(''), (_, _) => ShowMainWindow());
        menu.Items.Add("设置…", CreateMenuIcon(''), (_, _) => ShowSettingsWindow());
        menu.Items.Add("检查更新…", CreateMenuIcon(''), async (_, _) => await UpdateService.CheckAndPromptAsync(GetVisibleMainWindow(), manual: true));
        menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());

        var autoStartItem = new System.Windows.Forms.ToolStripMenuItem("开机自启", CreateMenuIcon(''))
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
        menu.Items.Add("退出", CreateMenuIcon(''), (_, _) => ExitApp());
        _trayIcon.ContextMenuStrip = menu;

        _trayIcon.DoubleClick += (_, _) => ShowMainWindow();

        // Event-driven (instant) via Windows audio COM notifications.
        // The "immediate" callback runs on the UI thread without debounce specifically
        // for lock enforcement so external default-device changes are reverted ASAP.
        _appRoutes = new AppRouteMonitor(Dispatcher, OnAppRoutesChanged);
        _appRoutes.SessionsChanged += () => AudioSessionsChanged?.Invoke();
        _deviceNotifier = new DeviceChangeNotifier(() =>
        {
            _appRoutes.DevicesChanged();
            CheckDeviceChanges();
        }, EnforceLockedProfileImmediately);

        // Low-frequency device safety net; app session discovery uses WASAPI notifications.
        // Verify routes of already-known sessions too: persisted per-app policy changes
        // do not have a complete public notification API. No process/session discovery here.
        _deviceWatchTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(5) };
        _deviceWatchTimer.Tick += (_, _) =>
        {
            CheckDeviceChanges();
            _appRoutes?.Schedule();
        };
        _deviceWatchTimer.Start();
        ProfileSchedules = new ProfileScheduleService(ApplyScheduledProfile,
            catchUpEnabled: () => SettingsService.Load().CatchUpProfileSchedules,
            verify: VerifyScheduledProfile);

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

        _ = MaybeAutoCheckForUpdatesAsync();
    }

    private Window? GetVisibleMainWindow() =>
        _mainWindow is { IsLoaded: true, IsVisible: true } ? _mainWindow : null;

    // WinForms' ToolStripMenuItem.Image needs an actual bitmap — unlike the WPF main-window
    // menu (MainWindow.xaml), which can render a Segoe MDL2 Assets glyph directly as a
    // TextBlock. Rasterizing the same glyph set here keeps the tray menu's icons consistent
    // with the main menu's.
    private static Bitmap CreateMenuIcon(char glyph)
    {
        var bmp = new Bitmap(16, 16);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
        using var font = new Font("Segoe MDL2 Assets", 11f);
        using var brush = new SolidBrush(Color.FromArgb(90, 90, 90));
        var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString(glyph.ToString(), font, brush, new RectangleF(0, 0, 16, 16), format);
        return bmp;
    }

    // Runs at most once per 24h (persisted via LastUpdateCheckUtc) so relaunching the app
    // repeatedly doesn't hammer the GitHub API. Silent unless a genuinely new, non-skipped
    // version is found — see UpdateService.CheckAndPromptAsync.
    private async Task MaybeAutoCheckForUpdatesAsync()
    {
        var settings = SettingsService.Load();
        if (!settings.AutoCheckUpdates) return;
        if (settings.LastUpdateCheckUtc is { } last && DateTime.UtcNow - last < TimeSpan.FromHours(24)) return;

        // Let the UI finish showing before doing network I/O.
        await Task.Delay(TimeSpan.FromSeconds(3));

        await UpdateService.CheckAndPromptAsync(GetVisibleMainWindow(), manual: false);
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
        if (result.SkippedAppNames.Count > 0)
        {
            var names = string.Join(", ", result.SkippedAppNames.Distinct(StringComparer.OrdinalIgnoreCase));
            body += $"\n\n{result.SkippedAppNames.Count} 个应用未运行，已跳过: {names}";
        }

        // If every device in the profile is offline (e.g. bluetooth disconnected), nothing was switched.
        bool hasPlayback = profile.PlaybackDeviceId != null;
        bool hasRecording = profile.RecordingDeviceId != null;
        int targetCount = (hasPlayback ? 1 : 0) + (hasRecording ? 1 : 0);
        bool allMissing = targetCount > 0 && result.MissingDeviceNames.Count == targetCount;

        string title;
        if (allMissing)
        {
            title = $"切换失败: {profile.Name}";
            body += $"\n\n以下设备未连接或未启用，未执行切换:\n{string.Join("\n", result.MissingDeviceNames)}";
        }
        else
        {
            title = $"已切换到 {profile.Name}";
            if (result.MissingDeviceNames.Count > 0)
                body += $"\n\n以下设备未连接或未启用，已跳过:\n{string.Join("\n", result.MissingDeviceNames)}";
        }

        switch (result.VoicemeeterStatus)
        {
            case VoicemeeterRestartStatus.Restarted:
                body += "\n\nVoicemeeter 音频引擎已重启";
                break;
            case VoicemeeterRestartStatus.NotRunning:
                body += "\n\nVoicemeeter 未运行，已跳过引擎重启";
                break;
            case VoicemeeterRestartStatus.NotInstalled:
                body += "\n\n未检测到 Voicemeeter 安装，已跳过引擎重启";
                break;
            case VoicemeeterRestartStatus.Failed:
                body += "\n\nVoicemeeter 引擎重启失败";
                break;
        }
        ToastService.Show(ToastService.TagProfileSwitch, title, body);
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
        _appRoutes?.Schedule();
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

        var settings = SettingsService.Load();
        bool inSuppressionWindowEarly = DateTime.Now < _suppressDeviceBalloonUntil;

        // Lock enforcement: if a profile is locked and the system has drifted away from it,
        // re-apply silently. Runs even on first poll so the lock survives across app restarts.
        if (!inSuppressionWindowEarly && settings.LockedProfileId is Guid lockedId)
        {
            var locked = ProfileService.GetAll().Find(p => p.Id == lockedId);
            if (locked != null && IsLockedDrifted(locked, playback, recording, playbackComm, recordingComm)
                && AreLockedDevicesAvailable(locked))
            {
                try
                {
                    ProfileApplyService.Apply(locked);
                    MarkOwnChange();
                    if (settings.NotifyDeviceChanged)
                        ToastService.Show(ToastService.TagDeviceChange,
                            $"已锁定到 {locked.Name}",
                            "外部切换已自动恢复");
                }
                catch { }
                // Re-read post-apply state for downstream logic.
                playback = AudioDeviceService.GetPlaybackDevices().Find(d => d.IsDefault);
                recording = AudioDeviceService.GetRecordingDevices().Find(d => d.IsDefault);
                playbackComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Render);
                recordingComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Capture);
            }
        }

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
    }

    private string ApplyScheduledProfile(Guid id)
    {
        if (SettingsService.Load().LockedProfileId.HasValue)
            return "已跳过：音频方案已锁定，请先解锁";
        var profile = ProfileService.GetAll().Find(p => p.Id == id);
        if (profile == null) return "已跳过：目标音频方案已删除";
        var result = ProfileApplyService.Apply(profile);
        MarkOwnChange();
        _mainWindow?.RefreshFromExternalChange();
        NotifyProfileApplied(profile, result);
        return $"等待确认：已提交「{profile.Name}」，正在核对设备和应用规则";
    }

    private (string Message, bool Pending) VerifyScheduledProfile(Guid id)
    {
        var profile = ProfileService.GetAll().Find(p => p.Id == id);
        if (profile == null) return ("检查已结束：目标音频方案已删除", false);
        if (ProfileApplyService.SelectedProfileId != id)
            return ("检查已结束：已切换到其他音频方案", false);
        var problems = new List<string>();
        var waiting = new List<string>();
        foreach (var flow in new[] { NAudio.CoreAudioApi.DataFlow.Render, NAudio.CoreAudioApi.DataFlow.Capture })
        {
            bool output = flow == NAudio.CoreAudioApi.DataFlow.Render;
            var devices = output ? AudioDeviceService.GetPlaybackDevices() : AudioDeviceService.GetRecordingDevices();
            var target = output ? profile.PlaybackDeviceId : profile.RecordingDeviceId;
            var system = devices.Find(d => d.IsDefault)?.Id;
            var label = output ? "输出" : "输入";
            if (target != null && !devices.Any(d => d.Id == target)) problems.Add(label + "设备未连接");
            else if (target != null && (system != target || AudioDeviceService.GetCommunicationsDefault(flow).Id != target))
                problems.Add(label + "默认设备不匹配");
            foreach (var rule in profile.AppOverrides)
            {
                var preset = AppProfileService.Get(rule.AppProfileId);
                if (preset == null) { problems.Add("音频预设已删除"); continue; }
                var name = System.IO.Path.GetFileName(rule.ExePath);
                var pids = (_appRoutes?.SessionSnapshot ?? []).Where(s => string.Equals(s.ExePath, rule.ExePath, StringComparison.OrdinalIgnoreCase))
                    .Select(s => s.ProcessId).Distinct().ToArray();
                if (pids.Length == 0) { waiting.Add(name + "：等待音频会话"); continue; }
                var expected = output ? preset.OutputDeviceId : preset.InputDeviceId;
                foreach (var pid in pids)
                {
                    var query = AppAudioRoutingService.QueryAppEndpoint(pid, flow);
                    var state = AppRouteStatus.Evaluate(true, query.Success, query.DeviceId, expected, system);
                    if (state != AppRouteState.Matched)
                    {
                        if (_appRoutes?.IsAwaitingConfirmation(rule.ExePath, flow) == true) waiting.Add(name + label + "：等待确认");
                        else problems.Add(name + label + (state == AppRouteState.Unknown ? "：无法读取路由" : "：设备不匹配"));
                    }
                }
            }
        }
        if (problems.Count > 0) return ("部分成功／需处理：" + string.Join("；", problems.Concat(waiting).Distinct()), false);
        if (waiting.Count > 0) return ("等待会话／确认：" + string.Join("；", waiting.Distinct()), true);
        return ("成功：默认设备及应用设备规则均已核对（未测试实际声音）", false);
    }

    private void OnAppRoutesChanged()
    {
        if (_appRoutes == null) return;
        var newlyDrifted = _appRoutes.Drifted.Except(_knownDriftedApps, StringComparer.OrdinalIgnoreCase).ToList();
        _knownDriftedApps = new(_appRoutes.Drifted, StringComparer.OrdinalIgnoreCase);
        if (newlyDrifted.Count > 0 && SettingsService.Load().NotifyAppDrift)
            ToastService.Show(ToastService.TagAppDrift, $"{newlyDrifted.Count} 个应用已偏离音频方案", string.Join(", ", newlyDrifted));
        _mainWindow?.RefreshFromExternalChange();
        AppRoutesChanged?.Invoke();
    }

    public IReadOnlyList<string> PendingAppRoutes => _appRoutes?.Pending ?? [];
    public bool IsAppRoutePending(string? path, NAudio.CoreAudioApi.DataFlow flow) =>
        _appRoutes?.IsAwaitingConfirmation(path, flow) == true;
    public event Action? AppRoutesChanged;
    public event Action? AudioSessionsChanged;
    public void RefreshAppRoutes() => _appRoutes?.Schedule();
    // True if a profile lock is active right now (any profile).
    public bool IsAnyProfileLocked() => SettingsService.Load().LockedProfileId.HasValue;

    // Returns true if a manual single-device change is allowed; refuses if any profile
    // is locked (the change would drift the lock and immediately revert anyway).
    public bool TryUserChangeDevice()
    {
        var settings = SettingsService.Load();
        if (settings.LockedProfileId is Guid lid)
        {
            var locked = ProfileService.GetAll().Find(p => p.Id == lid);
            ToastService.Show(ToastService.TagDeviceChange,
                "切换被阻止",
                $"已锁定到 {locked?.Name ?? "当前音频方案"}，请先解锁");
            return false;
        }
        return true;
    }

    // Returns true if the user-initiated apply is allowed; otherwise the lock is active
    // for a different profile and we refuse the switch outright (with a tray notification).
    public bool TryUserApplyProfile(DeviceProfile profile)
    {
        var settings = SettingsService.Load();
        if (settings.LockedProfileId is Guid lid && lid != profile.Id)
        {
            var locked = ProfileService.GetAll().Find(p => p.Id == lid);
            ToastService.Show(ToastService.TagProfileSwitch,
                "切换被阻止",
                $"已锁定到 {locked?.Name ?? "当前音频方案"}，请先解锁");
            return false;
        }
        return true;
    }

    // Fast-path called directly from IMMNotificationClient.OnDefaultDeviceChanged on the
    // UI thread. Reverts immediately when the locked profile drifts, no 150ms debounce.
    private void EnforceLockedProfileImmediately()
    {
        var settings = SettingsService.Load();
        if (settings.LockedProfileId is not Guid lid) return;
        if (DateTime.Now < _suppressDeviceBalloonUntil) return;

        var locked = ProfileService.GetAll().Find(p => p.Id == lid);
        if (locked == null) return;
        if (!AreLockedDevicesAvailable(locked)) return;

        var playback = AudioDeviceService.GetPlaybackDevices().Find(d => d.IsDefault);
        var recording = AudioDeviceService.GetRecordingDevices().Find(d => d.IsDefault);
        var playbackComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Render);
        var recordingComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Capture);
        if (!IsLockedDrifted(locked, playback, recording, playbackComm, recordingComm)) return;

        // Reserve the suppression window before Apply so that re-entrant
        // OnDefaultDeviceChanged events triggered by our own SetDefaultEndpoint won't
        // recurse into another full Apply cycle.
        _suppressDeviceBalloonUntil = DateTime.Now.AddMilliseconds(1500);
        _appRoutes?.Schedule();
        try
        {
            ProfileApplyService.Apply(locked);
            MarkOwnChange();
            if (settings.NotifyDeviceChanged)
                ToastService.Show(ToastService.TagDeviceChange,
                    $"已锁定到 {locked.Name}",
                    "外部切换已自动恢复");
        }
        catch { }
    }

    private static bool IsLockedDrifted(DeviceProfile p, AudioDeviceInfo? playback, AudioDeviceInfo? recording,
        (string? Id, string? Name) playbackComm, (string? Id, string? Name) recordingComm)
    {
        if (!string.IsNullOrEmpty(p.PlaybackDeviceId))
        {
            if (p.PlaybackDeviceId != playback?.Id) return true;
            if (p.PlaybackDeviceId != playbackComm.Id) return true;
        }
        if (!string.IsNullOrEmpty(p.RecordingDeviceId))
        {
            if (p.RecordingDeviceId != recording?.Id) return true;
            if (p.RecordingDeviceId != recordingComm.Id) return true;
        }
        return false;
    }

    private static bool AreLockedDevicesAvailable(DeviceProfile p)
    {
        if (!string.IsNullOrEmpty(p.PlaybackDeviceId))
        {
            var all = AudioDeviceService.GetPlaybackDevices();
            if (!all.Any(d => string.Equals(d.Id, p.PlaybackDeviceId, StringComparison.Ordinal))) return false;
        }
        if (!string.IsNullOrEmpty(p.RecordingDeviceId))
        {
            var all = AudioDeviceService.GetRecordingDevices();
            if (!all.Any(d => string.Equals(d.Id, p.RecordingDeviceId, StringComparison.Ordinal))) return false;
        }
        return true;
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

    public void ExitFromUI() => ExitApp();

    private void ExitApp()
    {
        ProfileSchedules?.Dispose();
        _deviceWatchTimer?.Stop();
        _deviceNotifier?.Dispose();
        _appRoutes?.Dispose();
        _trayIcon.Visible = false;
        _trayIcon.Dispose();
        _mainWindow?.ForceClose();
        _showSignal?.Dispose();
        _singleInstanceMutex?.ReleaseMutex();
        _singleInstanceMutex?.Dispose();
        VoicemeeterService.Shutdown();
        Shutdown();
    }
}
