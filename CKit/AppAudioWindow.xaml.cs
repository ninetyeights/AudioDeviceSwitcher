using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;
using NAudio.CoreAudioApi;

namespace AudioDeviceSwitcher;

public partial class AppAudioWindow : Window
{
    // Sentinel option meaning "follow system default" (cleared override).
    public record DeviceOption(string? Id, string Name);

    private static readonly DeviceOption FollowSystem = new(null, "跟随系统");

    public class SessionRow : INotifyPropertyChanged
    {
        public uint ProcessId { get; }
        public IReadOnlyList<string>? SessionInstanceKeys { get; }
        public string DisplayName { get; }
        public string SubText { get; }
        public ImageSource? Icon { get; }
        public List<DeviceOption> OutputOptions { get; }
        public List<DeviceOption> InputOptions { get; }

        private DeviceOption _selectedOutput = FollowSystem;
        public DeviceOption SelectedOutput
        {
            get => _selectedOutput;
            set
            {
                if (_selectedOutput == value) return;
                var old = _selectedOutput;
                _selectedOutput = value;
                OnChanged(nameof(SelectedOutput));
                OnChanged(nameof(IsOutputDrifted));
                OnChanged(nameof(IsDrifted));
                if (!_suppressApply) TryApply(DataFlow.Render, value.Id, old);
            }
        }

        private DeviceOption _selectedInput = FollowSystem;
        public DeviceOption SelectedInput
        {
            get => _selectedInput;
            set
            {
                if (_selectedInput == value) return;
                var old = _selectedInput;
                _selectedInput = value;
                OnChanged(nameof(SelectedInput));
                OnChanged(nameof(IsInputDrifted));
                OnChanged(nameof(IsDrifted));
                if (!_suppressApply) TryApply(DataFlow.Capture, value.Id, old);
            }
        }

        private double _volume = 1.0;
        public double Volume
        {
            get => _volume;
            set
            {
                if (Math.Abs(_volume - value) < 0.0001) return;
                _volume = value;
                OnChanged(nameof(Volume));
                OnChanged(nameof(VolumePercent));
                if (!_suppressApply) ApplyVolumeThrottled(value);
            }
        }

        public string VolumePercent => $"{(int)Math.Round(_volume * 100)}%";

        // Slider drag fires Volume many times a second; SetAppVolume re-enumerates every
        // audio device/session via COM, so applying on every tick still stutters even with
        // the PID cache below. Throttle instead of debounce: the first change in a burst
        // applies immediately (keeps it feeling real-time), further changes within the
        // window are coalesced and the last one applied on a trailing timer, capping the
        // actual COM call rate without ever going silent while dragging.
        private const int VolumeThrottleMs = 40;
        private long _lastVolumeApplyTicks = long.MinValue;
        private DispatcherTimer? _volumeTrailingTimer;
        private double _pendingVolume;

        private void ApplyVolumeThrottled(double value)
        {
            var now = Environment.TickCount64;
            if (now - _lastVolumeApplyTicks >= VolumeThrottleMs)
            {
                _volumeTrailingTimer?.Stop();
                _lastVolumeApplyTicks = now;
                DoApplyVolume(value);
                return;
            }

            _pendingVolume = value;
            if (_volumeTrailingTimer == null)
            {
                _volumeTrailingTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(VolumeThrottleMs) };
                _volumeTrailingTimer.Tick += (_, _) =>
                {
                    _volumeTrailingTimer!.Stop();
                    _lastVolumeApplyTicks = Environment.TickCount64;
                    DoApplyVolume(_pendingVolume);
                };
            }
            _volumeTrailingTimer.Stop();
            _volumeTrailingTimer.Start();
        }

        private void DoApplyVolume(double value)
        {
            foreach (var pid in AllPidsForThisApp())
                try { AudioSessionService.SetAppVolume(pid, (float)value, pid == ProcessId ? SessionInstanceKeys : null); } catch { }
        }

        private bool _isMuted;
        public bool IsMuted
        {
            get => _isMuted;
            set
            {
                if (_isMuted == value) return;
                _isMuted = value;
                OnChanged(nameof(IsMuted));
                if (!_suppressApply)
                    foreach (var pid in AllPidsForThisApp())
                        try { AudioSessionService.SetAppMute(pid, value, pid == ProcessId ? SessionInstanceKeys : null); } catch { }
            }
        }

        private bool _suppressApply;

        public string? ExecutablePath { get; }
        public string? ExpectedOutputId { get; set; }
        public string? ExpectedInputId { get; set; }
        public string? ExpectedProfileName { get; set; }
        public string? SystemDefaultOutputId { get; set; }
        public string? SystemDefaultInputId { get; set; }

        // "跟随系统" (Id == null) 时用系统默认设备 ID 比较，效果等价就不算偏离
        public bool IsOutputDrifted => ExpectedProfileName != null
            && !string.Equals(_selectedOutput.Id ?? SystemDefaultOutputId ?? "",
                ExpectedOutputId ?? "", StringComparison.OrdinalIgnoreCase);

        public bool IsInputDrifted => ExpectedProfileName != null
            && !string.Equals(_selectedInput.Id ?? SystemDefaultInputId ?? "",
                ExpectedInputId ?? "", StringComparison.OrdinalIgnoreCase);

        public bool IsDrifted => IsOutputDrifted || IsInputDrifted;

        public SessionRow(AppAudioSessionInfo info, List<DeviceOption> outputs, List<DeviceOption> inputs)
        {
            ProcessId = info.ProcessId;
            SessionInstanceKeys = info.SessionInstanceKeys;
            ExecutablePath = info.ExecutablePath;
            DisplayName = info.DisplayName;
            SubText = info.ExecutablePath is { Length: > 0 } p ? p : $"PID {info.ProcessId}";
            Icon = info.Icon;
            OutputOptions = outputs;
            InputOptions = inputs;
            // Captured straight off the session object during the same enumeration pass that
            // found this row (see AudioSessionService.EnumerateRawSessions) — avoids a second,
            // full device/session COM walk per row just to read the initial volume/mute state.
            _volume = info.Volume;
            _isMuted = info.Muted;
        }

        public void LoadCurrent()
        {
            _suppressApply = true;
            try
            {
                var outId = SafeGet(DataFlow.Render);
                var inId = SafeGet(DataFlow.Capture);
                _selectedOutput = OutputOptions.FirstOrDefault(o =>
                    string.Equals(o.Id, outId, StringComparison.OrdinalIgnoreCase)) ?? FollowSystem;
                _selectedInput = InputOptions.FirstOrDefault(o =>
                    string.Equals(o.Id, inId, StringComparison.OrdinalIgnoreCase)) ?? FollowSystem;

                OnChanged(nameof(SelectedOutput));
                OnChanged(nameof(SelectedInput));
                OnChanged(nameof(IsOutputDrifted));
                OnChanged(nameof(IsInputDrifted));
                OnChanged(nameof(IsDrifted));
                OnChanged(nameof(Volume));
                OnChanged(nameof(VolumePercent));
                OnChanged(nameof(IsMuted));
            }
            finally { _suppressApply = false; }
        }

        // Query across all PIDs for this exe (session PID + Process.MainModule matches);
        // return the first non-null value. Matches the PID set used when applying.
        private string? SafeGet(DataFlow flow)
        {
            foreach (var pid in AllPidsForThisApp())
            {
                try
                {
                    var id = AppAudioRoutingService.GetAppEndpoint(pid, flow);
                    if (!string.IsNullOrEmpty(id)) return id;
                }
                catch { }
            }
            return null;
        }

        // Process.GetProcesses() + per-process MainModule probing (inside
        // GetRunningPidsForExe) is a full-system scan and far too slow to redo on every
        // slider-drag tick or mute toggle. The set of PIDs sharing this exe is effectively
        // static for the lifetime of a row (rows are rebuilt wholesale on LoadSessions/
        // Refresh), so compute it once and cache it.
        private List<uint>? _pidCache;

        // Lets LoadSessions seed the cache from one batched system-wide scan (covering every
        // row's exe at once) instead of each row triggering its own Process.GetProcesses().
        public void PrimePidCache(Dictionary<string, List<uint>> pidsByExePath)
        {
            var list = new List<uint> { ProcessId };
            if (!string.IsNullOrEmpty(ExecutablePath) && pidsByExePath.TryGetValue(ExecutablePath, out var extra))
                foreach (var pid in extra)
                    if (pid != ProcessId) list.Add(pid);
            _pidCache = list;
        }

        private List<uint> AllPidsForThisApp()
        {
            if (_pidCache != null) return _pidCache;
            var list = new List<uint> { ProcessId };
            if (!string.IsNullOrEmpty(ExecutablePath))
                foreach (var pid in ProfileApplyService.GetRunningPidsForExe(ExecutablePath))
                    if (pid != ProcessId) list.Add(pid);
            return _pidCache = list;
        }

        private void TryApply(DataFlow flow, string? deviceId, DeviceOption previous)
        {
            try
            {
                foreach (var pid in AllPidsForThisApp())
                {
                    try { AppAudioRoutingService.SetAppEndpoint(pid, flow, deviceId); }
                    catch { /* best-effort per PID */ }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置失败：\n{ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                _suppressApply = true;
                try
                {
                    if (flow == DataFlow.Render) _selectedOutput = previous;
                    else _selectedInput = previous;
                    OnChanged(flow == DataFlow.Render ? nameof(SelectedOutput) : nameof(SelectedInput));
                }
                finally { _suppressApply = false; }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public AppAudioWindow()
    {
        InitializeComponent();
        Loaded += async (_, _) => await LoadSessionsAsync();
    }

    // BuildRows does nothing but COM/WMI/file I/O — none of it touches a UI element or the
    // Dispatcher (SessionRow is a plain INotifyPropertyChanged POCO with no subscribers until
    // it's bound below), so running it on a thread-pool thread via Task.Run genuinely frees
    // the UI thread instead of just reordering work on it. That's the difference between this
    // and the previous Dispatcher.BeginInvoke(Background) attempt: that still blocked the UI
    // thread for the full duration once its turn came up, so the window kept appearing to
    // freeze/white out while it ran. Off the UI thread, the window paints and stays
    // interactive immediately; the list just pops in a beat later.
    private async Task LoadSessionsAsync()
    {
        var rows = await Task.Run(BuildRows);

        // Kept on the UI thread deliberately: AppAudioRoutingService talks to an undocumented,
        // reverse-engineered WinRT COM factory (see its file header) with no documented
        // threading guarantee, unlike the well-worn WASAPI calls in BuildRows that this app
        // already exercises off-thread elsewhere without issue. Not worth the risk of a
        // cross-apartment COM failure for what's a handful of cheap per-row lookups anyway.
        foreach (var r in rows) r.LoadCurrent();

        SessionList.ItemsSource = rows;
        EmptyHint.Visibility = rows.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private static List<SessionRow> BuildRows()
    {
        var nicknames = SettingsService.Load().DeviceNicknames;
        string DisplayName(AudioDeviceInfo d) =>
            nicknames.TryGetValue(d.Id, out var nn) && !string.IsNullOrEmpty(nn)
                ? $"{nn} — {d.Name}" : d.Name;

        var playbackDevices = AudioDeviceService.GetPlaybackDevices();
        var recordingDevices = AudioDeviceService.GetRecordingDevices();

        var outputs = new List<DeviceOption> { FollowSystem };
        foreach (var d in playbackDevices)
            outputs.Add(new DeviceOption(d.Id, DisplayName(d)));

        var inputs = new List<DeviceOption> { FollowSystem };
        foreach (var d in recordingDevices)
            inputs.Add(new DeviceOption(d.Id, DisplayName(d)));

        var sessions = AudioSessionService.GetActiveAppSessions();
        var rows = sessions.Select(s => new SessionRow(s, outputs, inputs)).ToList();

        // Look up active system profile's AppOverrides — annotate each row with expected values.
        var currentPlayback = playbackDevices.Find(d => d.IsDefault);
        var currentRecording = recordingDevices.Find(d => d.IsDefault);
        foreach (var row in rows)
        {
            row.SystemDefaultOutputId = currentPlayback?.Id;
            row.SystemDefaultInputId = currentRecording?.Id;
        }
        var active = ProfileService.GetAll().Find(p =>
            p.PlaybackDeviceId == currentPlayback?.Id && p.RecordingDeviceId == currentRecording?.Id);
        if (active != null)
        {
            foreach (var row in rows)
            {
                if (string.IsNullOrEmpty(row.ExecutablePath)) continue;
                var ov = active.AppOverrides.Find(o =>
                    string.Equals(o.ExePath, row.ExecutablePath, StringComparison.OrdinalIgnoreCase));
                if (ov == null) continue;
                var ap = AppProfileService.Get(ov.AppProfileId);
                if (ap == null) continue;
                row.ExpectedOutputId = ap.OutputDeviceId;
                row.ExpectedInputId = ap.InputDeviceId;
                row.ExpectedProfileName = ap.Name;
            }
        }

        var exePaths = new HashSet<string>(
            rows.Select(r => r.ExecutablePath).Where(p => !string.IsNullOrEmpty(p))!,
            StringComparer.OrdinalIgnoreCase);
        var pidsByExePath = ProfileApplyService.GetRunningProcessesByPath(exePaths);
        foreach (var r in rows) r.PrimePidCache(pidsByExePath);

        return rows;
    }

    private async void Refresh_Click(object sender, RoutedEventArgs e) => await LoadSessionsAsync();

    private async void ClearAll_Click(object sender, RoutedEventArgs e)
    {
        var confirm = MessageBox.Show(
            "确定要清除所有应用的设备覆盖吗？所有应用将恢复使用系统默认设备。",
            "确认", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        try
        {
            AppAudioRoutingService.ClearAll();
            await LoadSessionsAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"清除失败：\n{ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
