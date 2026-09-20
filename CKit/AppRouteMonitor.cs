using System.IO;
using System.Windows.Threading;
using NAudio.CoreAudioApi;

namespace AudioDeviceSwitcher;

// UI-thread coordinator. Session discovery is event-driven; this timer is a one-shot
// settling/retry timer, never a recurring process/session scan.
public sealed class AppRouteMonitor : IDisposable
{
    private readonly AudioSessionNotifier _sessions;
    private readonly DispatcherTimer _settle;
    private readonly Action _changed;
    private readonly Dictionary<(string Path, DataFlow Flow), int> _pending = new();
    private readonly Dictionary<(uint Pid, DataFlow Flow), DateTime> _lastRepair = new();
    private Guid? _profileId;
    public HashSet<string> Drifted { get; private set; } = new(StringComparer.OrdinalIgnoreCase);
    public List<string> Pending { get; private set; } = [];
    public event Action? SessionsChanged;
    public IReadOnlyList<RoutingSession> SessionSnapshot => _sessions.Snapshot;
    public bool IsAwaitingConfirmation(string? path, DataFlow flow) => path != null
        && _pending.TryGetValue((path.ToUpperInvariant(), flow), out var attempts) && attempts < 3;

    public AppRouteMonitor(Dispatcher dispatcher, Action changed)
    {
        _changed = changed;
        _settle = new DispatcherTimer(DispatcherPriority.Background, dispatcher)
            { Interval = TimeSpan.FromSeconds(1) };
        _settle.Tick += (_, _) => { _settle.Stop(); Check(); };
        _sessions = new AudioSessionNotifier(dispatcher, () =>
        {
            Schedule();
            SessionsChanged?.Invoke();
        });
        ProfileApplyService.Applying += OnApplying;
    }

    public void DevicesChanged() { _sessions.RefreshDevices(); Schedule(); }
    public void Schedule()
    {
        // Do not indefinitely postpone verification when many sessions arrive together.
        if (!_settle.IsEnabled) _settle.Start();
    }

    private void OnApplying(DeviceProfile profile)
    {
        Arm(profile);
        Drifted.Clear();
        Pending = profile.AppOverrides.Select(o => $"{Path.GetFileName(o.ExePath)}：等待确认").ToList();
        _changed();
        _settle.Stop();
        Schedule();
    }

    private void Arm(DeviceProfile profile)
    {
        _profileId = profile.Id;
        _pending.Clear();
        _lastRepair.Clear();
        foreach (var ov in profile.AppOverrides)
            foreach (var flow in new[] { DataFlow.Render, DataFlow.Capture })
                _pending[(ov.ExePath.ToUpperInvariant(), flow)] = 0;
    }

    private void Check()
    {
        var playback = AudioDeviceService.GetPlaybackDevices().Find(d => d.IsDefault)?.Id;
        var recording = AudioDeviceService.GetRecordingDevices().Find(d => d.IsDefault)?.Id;
        var profile = ProfileApplyService.FindActiveProfile(playback, recording);
        var drifted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pending = new List<string>();
        if (profile == null)
        {
            _profileId = null;
            _pending.Clear();
        }
        else
        {
            if (_profileId != profile.Id) Arm(profile);
            var snapshot = _sessions.Snapshot;
            foreach (var ov in profile.AppOverrides)
            {
                var ap = AppProfileService.Get(ov.AppProfileId);
                if (ap == null) continue;
                var name = Path.GetFileName(ov.ExePath);
                // A session identifies the app's process; persisted input and output
                // preferences can both be queried/set through that PID. A capture
                // stream is not required to configure the microphone route.
                var pids = snapshot.Where(s =>
                    string.Equals(s.ExePath, ov.ExePath, StringComparison.OrdinalIgnoreCase))
                    .Select(s => s.ProcessId).Distinct().ToArray();
                foreach (var flow in new[] { DataFlow.Render, DataFlow.Capture })
                {
                    var expected = flow == DataFlow.Render ? ap.OutputDeviceId : ap.InputDeviceId;
                    var system = flow == DataFlow.Render ? playback : recording;
                    var label = flow == DataFlow.Render ? "输出" : "输入";
                    var key = (ov.ExePath.ToUpperInvariant(), flow);
                    if (pids.Length == 0)
                    {
                        if (expected != null) pending.Add($"{name}（{label}）：等待音频会话");
                        continue;
                    }
                    bool allMatched = true;
                    foreach (var pid in pids)
                    {
                        var query = AppAudioRoutingService.QueryAppEndpoint(pid, flow);
                        var state = AppRouteStatus.Evaluate(true, query.Success, query.DeviceId, expected, system);
                        if (state == AppRouteState.Matched) continue;
                        allMatched = false;
                        bool waiting = _pending.TryGetValue(key, out var attempts) && attempts < 3;
                        bool locked = SettingsService.Load().LockedProfileId == profile.Id;
                        var repairKey = (pid, flow);
                        bool canRepair = !_lastRepair.TryGetValue(repairKey, out var last) ||
                            DateTime.UtcNow - last >= TimeSpan.FromSeconds(5);
                        if (waiting || (locked && state == AppRouteState.Drifted && canRepair))
                        {
                            // A repair attempt is not evidence of success. Preserve confirmed
                            // drift until a later read matches, avoiding repeated "new" alerts.
                            if (!waiting && state == AppRouteState.Drifted) drifted.Add(name);
                            try { AppAudioRoutingService.SetAppEndpoint(pid, flow, expected); }
                            catch { /* Verification below distinguishes unavailable from mismatched. */ }
                            _lastRepair[repairKey] = DateTime.UtcNow;
                            pending.Add($"{name}（{label}）：等待确认");
                            Schedule();
                        }
                        else if (state == AppRouteState.Unknown)
                            pending.Add($"{name}（{label}）：无法读取路由");
                        else drifted.Add(name);
                    }
                    if (allMatched) _pending.Remove(key);
                    else if (_pending.TryGetValue(key, out var attempts)) _pending[key] = attempts + 1;
                }
            }
        }
        Drifted = drifted;
        Pending = pending.Distinct().ToList();
        _changed();
    }

    public void Dispose()
    {
        ProfileApplyService.Applying -= OnApplying;
        _settle.Stop();
        _sessions.Dispose();
    }
}
