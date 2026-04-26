using System.Windows;
using System.Windows.Threading;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;

namespace AudioDeviceSwitcher;

// Listens to Windows audio-device COM notifications (IMMNotificationClient) and fires
// a single debounced callback back on the UI thread. Replaces polling for instant reaction
// to default-device changes, device add/remove, and enable/disable.
public class DeviceChangeNotifier : IMMNotificationClient, IDisposable
{
    private readonly MMDeviceEnumerator _enumerator = new();
    private readonly Action _onChange;
    private readonly Action? _onDefaultChangedImmediate;
    private readonly DispatcherTimer _debounce;
    private bool _registered;

    public DeviceChangeNotifier(Action onChange, Action? onDefaultChangedImmediate = null)
    {
        _onChange = onChange;
        _onDefaultChangedImmediate = onDefaultChangedImmediate;
        _debounce = new DispatcherTimer(DispatcherPriority.Background, Application.Current.Dispatcher)
        {
            Interval = TimeSpan.FromMilliseconds(150),
        };
        _debounce.Tick += (_, _) =>
        {
            _debounce.Stop();
            try { _onChange(); } catch { }
        };
        _enumerator.RegisterEndpointNotificationCallback(this);
        _registered = true;
    }

    // IMMNotificationClient — called on COM/MTA threads; marshal to UI thread via dispatcher timer.
    public void OnDeviceStateChanged(string deviceId, DeviceState newState) => Kick();
    public void OnDeviceAdded(string pwstrDeviceId) => Kick();
    public void OnDeviceRemoved(string deviceId) => Kick();
    public void OnDefaultDeviceChanged(DataFlow flow, Role role, string defaultDeviceId)
    {
        // Fire the lock-enforcement callback first, with no debounce — this lets a locked
        // profile snap the default back before the user perceives the switch in mmsys.cpl.
        if (_onDefaultChangedImmediate != null)
        {
            Application.Current.Dispatcher.BeginInvoke(_onDefaultChangedImmediate, DispatcherPriority.Send);
        }
        Kick();
    }
    public void OnPropertyValueChanged(string pwstrDeviceId, PropertyKey key)
    {
        // Ignored — very noisy (fires on volume, format, etc.). We only care about structural changes.
    }

    private void Kick()
    {
        // Restart the debounce on the UI thread; multiple COM events within 150ms collapse to one refresh.
        Application.Current.Dispatcher.BeginInvoke(() =>
        {
            _debounce.Stop();
            _debounce.Start();
        });
    }

    public void Dispose()
    {
        if (_registered)
        {
            try { _enumerator.UnregisterEndpointNotificationCallback(this); } catch { }
            _registered = false;
        }
        _debounce.Stop();
        _enumerator.Dispose();
    }
}
