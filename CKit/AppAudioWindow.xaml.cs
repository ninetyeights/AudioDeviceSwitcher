using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
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
            ExecutablePath = info.ExecutablePath;
            DisplayName = info.DisplayName;
            SubText = info.ExecutablePath is { Length: > 0 } p ? p : $"PID {info.ProcessId}";
            Icon = info.Icon;
            OutputOptions = outputs;
            InputOptions = inputs;
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

        private IEnumerable<uint> AllPidsForThisApp()
        {
            yield return ProcessId;
            if (string.IsNullOrEmpty(ExecutablePath)) yield break;
            foreach (var pid in ProfileApplyService.GetRunningPidsForExe(ExecutablePath))
                if (pid != ProcessId) yield return pid;
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
        Loaded += (_, _) => LoadSessions();
    }

    private void LoadSessions()
    {
        var nicknames = SettingsService.Load().DeviceNicknames;
        string DisplayName(AudioDeviceInfo d) =>
            nicknames.TryGetValue(d.Id, out var nn) && !string.IsNullOrEmpty(nn)
                ? $"{nn} — {d.Name}" : d.Name;

        var outputs = new List<DeviceOption> { FollowSystem };
        foreach (var d in AudioDeviceService.GetPlaybackDevices())
            outputs.Add(new DeviceOption(d.Id, DisplayName(d)));

        var inputs = new List<DeviceOption> { FollowSystem };
        foreach (var d in AudioDeviceService.GetRecordingDevices())
            inputs.Add(new DeviceOption(d.Id, DisplayName(d)));

        var sessions = AudioSessionService.GetActiveAppSessions();
        var rows = sessions.Select(s => new SessionRow(s, outputs, inputs)).ToList();

        // Look up active system profile's AppOverrides — annotate each row with expected values.
        var currentPlayback = AudioDeviceService.GetPlaybackDevices().Find(d => d.IsDefault);
        var currentRecording = AudioDeviceService.GetRecordingDevices().Find(d => d.IsDefault);
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

        foreach (var r in rows) r.LoadCurrent();

        SessionList.ItemsSource = rows;
        EmptyHint.Visibility = rows.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void Refresh_Click(object sender, RoutedEventArgs e) => LoadSessions();

    private void ClearAll_Click(object sender, RoutedEventArgs e)
    {
        var confirm = MessageBox.Show(
            "确定要清除所有应用的设备覆盖吗？所有应用将恢复使用系统默认设备。",
            "确认", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        try
        {
            AppAudioRoutingService.ClearAll();
            LoadSessions();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"清除失败：\n{ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
