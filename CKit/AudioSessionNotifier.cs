using System.Collections.Concurrent;
using System.Diagnostics;
using System.Windows.Threading;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;

namespace AudioDeviceSwitcher;

public record RoutingSession(uint ProcessId, string ExePath, DataFlow Flow);

// Own all WASAPI registrations on a long-lived MTA thread. Callbacks only queue work;
// neither COM callback threads nor the UI thread enumerate devices or unregister clients.
public sealed class AudioSessionNotifier : IDisposable
{
    private readonly BlockingCollection<Action> _work = new();
    private readonly Thread _thread;
    private readonly Dispatcher _dispatcher;
    private readonly Action _changed;
    private readonly Dictionary<string, Endpoint> _endpoints = new();
    private RoutingSession[] _snapshot = [];
    private int _disposed;
    public RoutingSession[] Snapshot => Volatile.Read(ref _snapshot);

    public AudioSessionNotifier(Dispatcher dispatcher, Action changed)
    {
        _dispatcher = dispatcher;
        _changed = changed;
        _thread = new Thread(Run) { IsBackground = true, Name = "Audio session notifications" };
        _thread.SetApartmentState(ApartmentState.MTA);
        _thread.Start();
        RefreshDevices();
    }

    private void Queue(Action action)
    {
        if (Volatile.Read(ref _disposed) != 0) return;
        try { _work.Add(action); } catch (InvalidOperationException) { }
    }

    private void Run()
    {
        try
        {
            foreach (var action in _work.GetConsumingEnumerable())
                try { action(); } catch (Exception ex) { Debug.WriteLine(ex); }
        }
        finally
        {
            foreach (var endpoint in _endpoints.Values) endpoint.Dispose();
            _work.Dispose();
        }
    }

    public void RefreshDevices() => Queue(() =>
    {
        using var enumerator = new MMDeviceEnumerator();
        var live = new HashSet<string>();
        foreach (var device in enumerator.EnumerateAudioEndPoints(DataFlow.All, DeviceState.Active))
        {
            var id = device.ID;
            live.Add(id);
            if (_endpoints.ContainsKey(id)) { device.Dispose(); continue; }
            Endpoint? endpoint = null;
            try
            {
                endpoint = new Endpoint(this, device);
                _endpoints.Add(id, endpoint);
                endpoint.Initialize();
            }
            catch
            {
                _endpoints.Remove(id);
                if (endpoint != null) endpoint.Dispose(); else device.Dispose();
            }
        }
        foreach (var id in _endpoints.Keys.Where(id => !live.Contains(id)).ToArray())
        {
            _endpoints[id].Dispose();
            _endpoints.Remove(id);
        }
        Publish();
    });

    private void Publish()
    {
        Volatile.Write(ref _snapshot, _endpoints.Values.SelectMany(e => e.Sessions.Values)
            .Select(s => s.Info).Distinct().ToArray());
        _dispatcher.BeginInvoke(() => { if (_disposed == 0) _changed(); });
    }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        _work.CompleteAdding();
        // No UI invocation is awaited by the worker, so shutdown cannot deadlock.
        _thread.Join();
    }

    private sealed class Endpoint(AudioSessionNotifier owner, MMDevice device) : IDisposable
    {
        private AudioSessionManager? _manager;
        private bool _disposed;
        public Dictionary<string, Session> Sessions { get; } = new();

        public void Initialize()
        {
            _manager = device.AudioSessionManager;
            _manager.OnSessionCreated += Created;
            // GetCount is required by WASAPI to enable session notifications.
            var existing = _manager.Sessions;
            for (int i = 0; i < existing.Count; i++) Add(existing[i]);
        }

        private void Created(object sender, IAudioSessionControl control) => owner.Queue(() =>
        {
            Add(new AudioSessionControl(control));
            owner.Publish();
        });

        private void Add(AudioSessionControl control)
        {
            bool retained = false;
            try
            {
                if (_disposed || control.State == AudioSessionState.AudioSessionStateExpired) return;
                var pid = control.GetProcessID;
                if (pid == 0) return;
                var key = control.GetSessionInstanceIdentifier;
                if (Sessions.ContainsKey(key)) return;
                using var process = Process.GetProcessById((int)pid);
                string? path = null;
                try { path = process.MainModule?.FileName; } catch { }
                path ??= AudioSessionService.GetProcessImagePath(pid);
                if (string.IsNullOrEmpty(path)) return;
                var session = new Session(control, new(pid, path, device.DataFlow), () => owner.Queue(() =>
                {
                    if (_disposed) return;
                    if (control.State == AudioSessionState.AudioSessionStateExpired)
                    {
                        Sessions.Remove(key);
                        control.Dispose();
                    }
                    owner.Publish();
                }), () => owner.Queue(() =>
                {
                    if (_disposed) return;
                    Sessions.Remove(key);
                    control.Dispose();
                    owner.Publish();
                }));
                control.RegisterEventClient(session);
                Sessions.Add(key, session);
                retained = true;
            }
            catch (Exception ex) { Debug.WriteLine(ex); }
            finally { if (!retained) control.Dispose(); }
        }

        public void Dispose()
        {
            _disposed = true;
            foreach (var session in Sessions.Values) try { session.Control.Dispose(); } catch { }
            Sessions.Clear();
            if (_manager != null)
            {
                _manager.OnSessionCreated -= Created;
                try { _manager.Dispose(); } catch { }
            }
            device.Dispose();
        }
    }

    private sealed class Session(AudioSessionControl control, RoutingSession info,
        Action changed, Action disconnected) : IAudioSessionEventsHandler
    {
        public AudioSessionControl Control => control;
        public RoutingSession Info => info;
        public void OnStateChanged(AudioSessionState state) => changed();
        public void OnSessionDisconnected(AudioSessionDisconnectReason reason) => disconnected();
        public void OnVolumeChanged(float volume, bool isMuted) { }
        public void OnDisplayNameChanged(string displayName) { }
        public void OnIconPathChanged(string iconPath) { }
        public void OnChannelVolumeChanged(uint channelCount, IntPtr newVolumes, uint channelIndex) { }
        public void OnGroupingParamChanged(ref Guid groupingId) { }
    }
}
