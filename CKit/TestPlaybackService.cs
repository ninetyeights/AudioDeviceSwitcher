using NAudio.CoreAudioApi;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;

namespace AudioDeviceSwitcher;

// 在指定播放设备上播放一段测试音（不依赖系统默认设备）。
// 同一时间只播放一个，再次触发会停止上一个。
public static class TestPlaybackService
{
    private static readonly object _gate = new();
    private static WasapiOut? _current;
    private static MMDevice? _currentDevice;

    public static void Play(string deviceId)
    {
        Stop();

        MMDevice? device = null;
        WasapiOut? output = null;
        try
        {
            using var enumerator = new MMDeviceEnumerator();
            device = enumerator.GetDevice(deviceId);
            if (device == null || device.State != DeviceState.Active)
            {
                device?.Dispose();
                throw new InvalidOperationException("设备不可用");
            }

            // 1 秒 440Hz 正弦音，双声道，音量适中。
            var tone = new SignalGenerator(44100, 2)
            {
                Gain = 0.25,
                Frequency = 440,
                Type = SignalGeneratorType.Sin,
            }.Take(TimeSpan.FromSeconds(1));

            output = new WasapiOut(device, AudioClientShareMode.Shared, false, 100);
            output.Init(tone);
            output.PlaybackStopped += (_, _) =>
            {
                lock (_gate)
                {
                    if (!ReferenceEquals(_current, output)) return;
                    DisposeCurrentNoLock();
                }
            };

            lock (_gate)
            {
                _current = output;
                _currentDevice = device;
                output.Play();
            }
        }
        catch
        {
            try { output?.Dispose(); } catch { }
            try { device?.Dispose(); } catch { }
            throw;
        }
    }

    public static void Stop()
    {
        lock (_gate)
        {
            if (_current == null) return;
            try { _current.Stop(); } catch { }
            DisposeCurrentNoLock();
        }
    }

    private static void DisposeCurrentNoLock()
    {
        try { _current?.Dispose(); } catch { }
        try { _currentDevice?.Dispose(); } catch { }
        _current = null;
        _currentDevice = null;
    }
}
