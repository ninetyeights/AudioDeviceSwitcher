using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using NAudio.CoreAudioApi;
using NAudio.Wave;

namespace AudioDeviceSwitcher;

public partial class MainWindow : Window
{
    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);
    private const int VK_LBUTTON = 0x01;
    private static bool IsLeftMouseDown() => (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOZORDER = 0x0004;
    private const uint SWP_NOACTIVATE = 0x0010;

    private MiniWindow? _miniWindow;
    private HotkeyService? _hotkeyService;
    private bool _forceClose;
    private string _lastStateSignature = "";
    private string? _selectedPlaybackId;
    private string? _selectedRecordingId;
    private System.Windows.Threading.DispatcherTimer? _peakTimer;
    private MMDevice? _peakPlaybackDevice;
    private MMDevice? _peakRecordingDevice;
    // Keepalive capture: AudioMeterInformation on capture endpoints only updates while
    // some client is actually recording. We open a shared-mode capture and discard the
    // bytes so the meter stays active even when no other app is using the mic.
    private WasapiCapture? _peakRecordingCapture;
    private FrameworkElement? _peakPlaybackTrack;
    private FrameworkElement? _peakPlaybackFill;
    private FrameworkElement? _peakRecordingTrack;
    private FrameworkElement? _peakRecordingFill;
    private double _peakPlaybackDisplayed;
    private double _peakRecordingDisplayed;
    private FrameworkElement[]? _vmOutputTracks;
    private FrameworkElement[]? _vmOutputFills;
    private FrameworkElement[]? _vmInputTracks;
    private FrameworkElement[]? _vmInputFills;
    private double[]? _vmOutputDisplayed;
    private double[]? _vmInputDisplayed;
    private int _vmOutCount;
    private int _vmInCount;
    private Point _profileDragStart;
    private Guid _profileDragSourceId;
    // Per-strip "user just clicked mute" overrides — the dirty poll may read a stale
    // state right after our write, so we trust this for a short window.
    private readonly Dictionary<int, (bool Muted, DateTime When)> _pendingMuteWrites = new();
    private static readonly TimeSpan PendingMuteWindow = TimeSpan.FromSeconds(2);
    private VoicemeeterIoState? _lastVoicemeeterState;

    public MainWindow()
    {
        InitializeComponent();
        ApplySettings();
        SourceInitialized += (_, _) =>
        {
            _hotkeyService = new HotkeyService(this);
            RegisterProfileHotkeys();

            // Restore position in physical pixels via Win32 — works correctly across
            // monitors with different DPI (WPF Left/Top cannot be used reliably here).
            var s = SettingsService.Load();
            if (s.MainWindowLeft is double l && s.MainWindowTop is double t)
            {
                var hwnd = new WindowInteropHelper(this).Handle;
                if (hwnd != IntPtr.Zero)
                    SetWindowPos(hwnd, IntPtr.Zero, (int)l, (int)t, 0, 0,
                        SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
            }
        };
        LoadProfiles();
        LoadDevices();
        Closing += (_, e) =>
        {
            SaveSettings();
            if (!_forceClose)
            {
                e.Cancel = true;
                Hide();
            }
        };
        // Only save during user drags (mouse held). DPI-adjust LocationChanged fires
        // without mouse pressed, so it's naturally filtered.
        LocationChanged += (_, _) =>
        {
            if (IsLeftMouseDown()) SaveSettings();
        };
        SizeChanged += (_, _) =>
        {
            if (IsLeftMouseDown()) SaveSettings();
        };
        IsVisibleChanged += (_, _) =>
        {
            if (IsVisible)
            {
                _lastStateSignature = "";
                RefreshFromExternalChange();
                StartVoicemeeterPolling();
                if (WindowState != WindowState.Minimized) StartPeakMetering();
            }
            else
            {
                StopVoicemeeterPolling();
                StopPeakMetering();
            }
        };
        StateChanged += (_, _) =>
        {
            if (!IsVisible) return;
            if (WindowState == WindowState.Minimized) StopPeakMetering();
            else StartPeakMetering();
        };
    }

    private void StartPeakMetering()
    {
        if (_peakTimer != null) return;
        if (_peakRecordingCapture == null) RefreshPeakDevices();
        _peakTimer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(50),
        };
        _peakTimer.Tick += PeakTimer_Tick;
        _peakTimer.Start();
    }

    private void StopPeakMetering()
    {
        _peakTimer?.Stop();
        _peakTimer = null;
        _peakPlaybackDisplayed = 0;
        _peakRecordingDisplayed = 0;
        StopRecordingKeepalive();
    }

    private void RefreshPeakDevices()
    {
        StopRecordingKeepalive();
        try { _peakPlaybackDevice?.Dispose(); } catch { }
        try { _peakRecordingDevice?.Dispose(); } catch { }
        _peakPlaybackDevice = null;
        _peakRecordingDevice = null;
        try
        {
            using var enumerator = new MMDeviceEnumerator();
            try { _peakPlaybackDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia); } catch { }
            try { _peakRecordingDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Multimedia); } catch { }
        }
        catch { }

        if (_peakRecordingDevice != null)
        {
            try
            {
                var cap = new WasapiCapture(_peakRecordingDevice) { ShareMode = AudioClientShareMode.Shared };
                cap.DataAvailable += (_, _) => { };
                cap.StartRecording();
                _peakRecordingCapture = cap;
            }
            catch
            {
                try { _peakRecordingCapture?.Dispose(); } catch { }
                _peakRecordingCapture = null;
            }
        }
    }

    private void StopRecordingKeepalive()
    {
        if (_peakRecordingCapture == null) return;
        try { _peakRecordingCapture.StopRecording(); } catch { }
        try { _peakRecordingCapture.Dispose(); } catch { }
        _peakRecordingCapture = null;
    }

    private void PeakTimer_Tick(object? sender, EventArgs e)
    {
        UpdatePeak(_peakPlaybackDevice, _peakPlaybackTrack, _peakPlaybackFill, ref _peakPlaybackDisplayed);
        UpdatePeak(_peakRecordingDevice, _peakRecordingTrack, _peakRecordingFill, ref _peakRecordingDisplayed);
        UpdateVoicemeeterPeaks();
    }

    private void UpdateVoicemeeterPeaks()
    {
        if (_vmOutCount == 0 && _vmInCount == 0) return;
        if (_vmOutputTracks == null || _vmOutputFills == null
            || _vmInputTracks == null || _vmInputFills == null
            || _vmOutputDisplayed == null || _vmInputDisplayed == null) return;
        var (outs, ins) = VoicemeeterService.GetPeakLevels(_vmOutCount, _vmInCount);
        for (int i = 0; i < outs.Length && i < _vmOutputTracks.Length; i++)
            UpdateVmBar(_vmOutputTracks[i], _vmOutputFills[i], outs[i], ref _vmOutputDisplayed[i]);
        for (int i = 0; i < ins.Length && i < _vmInputTracks.Length; i++)
            UpdateVmBar(_vmInputTracks[i], _vmInputFills[i], ins[i], ref _vmInputDisplayed[i]);
    }

    private static void UpdateVmBar(FrameworkElement track, FrameworkElement fill, float peak, ref double displayed)
    {
        double clamped = Math.Min(1.0, peak);
        displayed = clamped >= displayed ? clamped : Math.Max(clamped, displayed * 0.85);
        double w = track.ActualWidth * displayed;
        if (double.IsNaN(w) || w < 0) w = 0;
        fill.Width = w;
    }

    private static void UpdatePeak(MMDevice? device, FrameworkElement? track, FrameworkElement? fill, ref double displayed)
    {
        if (track == null || fill == null) return;
        float peak = 0f;
        if (device != null)
        {
            try { peak = device.AudioMeterInformation.MasterPeakValue; } catch { peak = 0f; }
        }
        // Visual smoothing: rises instantly, decays gradually.
        displayed = peak >= displayed ? peak : Math.Max(peak, displayed * 0.85);
        double w = track.ActualWidth * displayed;
        if (double.IsNaN(w) || w < 0) w = 0;
        fill.Width = w;
    }

    // Background thread tight-poll: 50ms when locked (so external Voicemeeter UI changes
    // are reverted in <100ms), 1.5s otherwise. IsParametersDirty is a microsecond-cheap
    // memory read so this is essentially free.
    private CancellationTokenSource? _voicemeeterPollCts;
    private Thread? _voicemeeterPollThread;

    private void StartVoicemeeterPolling()
    {
        if (_voicemeeterPollThread != null) return;
        _voicemeeterPollCts = new CancellationTokenSource();
        var token = _voicemeeterPollCts.Token;
        _voicemeeterPollThread = new Thread(() => VoicemeeterPollLoop(token))
        {
            IsBackground = true,
            Name = "VoicemeeterPoll",
        };
        _voicemeeterPollThread.Start();
    }

    private void StopVoicemeeterPolling()
    {
        try { _voicemeeterPollCts?.Cancel(); } catch { }
        _voicemeeterPollCts?.Dispose();
        _voicemeeterPollCts = null;
        _voicemeeterPollThread = null;
    }

    private void VoicemeeterPollLoop(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            try
            {
                if (VoicemeeterService.IsParametersDirty())
                {
                    var state = VoicemeeterService.GetIoState();
                    if (state != null) state = EnforceVoicemeeterLock(state);
                    Dispatcher.BeginInvoke(() => RenderVoicemeeterStatus(state));
                }
            }
            catch { }
            int interval = SettingsService.Load().VoicemeeterMuteLocked ? 50 : 1500;
            if (token.WaitHandle.WaitOne(interval)) break;
        }
    }

    private void ApplySettings()
    {
        // Position is restored in SourceInitialized via Win32 (physical pixels).
        // Only size is restored here — it's in DIPs and behaves consistently.
        var s = SettingsService.Load();
        WindowStartupLocation = WindowStartupLocation.Manual;
        if (s.MainWindowWidth is double w && w >= MinWidth) Width = w;
        if (s.MainWindowHeight is double h && h >= MinHeight) Height = h;
        ShowHiddenCheck.IsChecked = s.ShowHiddenDevices;
        ShowDisabledCheck.IsChecked = s.ShowDisabledDevices;
    }

    private void SaveSettings()
    {
        if (WindowState != WindowState.Normal) return;

        var s = SettingsService.Load();
        // Save position in physical pixels via Win32 — bypasses WPF DIP ambiguity
        // across multi-DPI monitors.
        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out var rect))
        {
            s.MainWindowLeft = rect.Left;
            s.MainWindowTop = rect.Top;
        }
        if (IsValid(Width)) s.MainWindowWidth = Width;
        if (IsValid(Height)) s.MainWindowHeight = Height;
        SettingsService.Save();
    }

    private static bool IsValid(double v) => !double.IsNaN(v) && !double.IsInfinity(v);

    private static bool IsOnScreen(double left, double top)
    {
        // Virtual screen bounds in DIPs (matches WPF Left/Top units under PerMonitorV2).
        var vLeft = SystemParameters.VirtualScreenLeft;
        var vTop = SystemParameters.VirtualScreenTop;
        var vRight = vLeft + SystemParameters.VirtualScreenWidth;
        var vBottom = vTop + SystemParameters.VirtualScreenHeight;
        return left >= vLeft - 50 && left <= vRight - 50
            && top >= vTop - 10 && top <= vBottom - 50;
    }

    public void OpenMiniWindow()
    {
        if (_miniWindow is { IsLoaded: true })
        {
            _miniWindow.Activate();
            return;
        }

        _miniWindow = new MiniWindow(() =>
        {
            LoadDevices();
            LoadProfiles();
        });
        _miniWindow.Closed += (_, _) =>
        {
            _miniWindow = null;
            UpdateMiniMenuCheck();
        };
        _miniWindow.Show();

        var settings = SettingsService.Load();
        settings.MiniWindowVisible = true;
        SettingsService.Save();

        UpdateMiniMenuCheck();
    }

    private void UpdateMiniMenuCheck()
    {
        if (MiniWindowMenuItem == null) return;
        bool open = _miniWindow is { IsLoaded: true };
        MiniWindowMenuItem.IsChecked = open;
        MiniWindowMenuItem.Header = open ? "\u5173\u95ED\u8FF7\u4F60\u7A97\u53E3" : "\u6253\u5F00\u8FF7\u4F60\u7A97\u53E3";
    }

    public void ForceClose()
    {
        _forceClose = true;
        _hotkeyService?.Dispose();
        _miniWindow?.Close();
        Close();
    }

    public void RefreshFromExternalChange()
    {
        // Nothing to refresh if both windows are hidden/closed
        if (!IsVisible && _miniWindow == null) return;

        // Skip if state hasn't changed (avoid needless UI rebuild)
        var signature = ComputeStateSignature();
        if (signature == _lastStateSignature) return;
        _lastStateSignature = signature;

        if (IsVisible)
        {
            LoadDevices();
            LoadProfiles();
        }
        _miniWindow?.LoadProfiles();
    }

    private static string ComputeStateSignature()
    {
        var allPlayback = AudioDeviceService.GetPlaybackDevices(true);
        var allRecording = AudioDeviceService.GetRecordingDevices(true);
        var playback = allPlayback.Find(d => d.IsDefault);
        var recording = allRecording.Find(d => d.IsDefault);
        var playbackComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Render);
        var recordingComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Capture);
        var bt = AudioDeviceService.HasBluetoothDevice();
        var profileCount = ProfileService.GetAll().Count;
        var drift = string.Join(",", ((App)Application.Current).DriftedApps);
        // Fingerprint every device's id+state so external enable/disable/add/remove triggers refresh.
        var deviceFp = string.Join(",",
            allPlayback.Concat(allRecording).Select(d => $"{d.Id}:{(d.IsDisabled ? "D" : "A")}"));
        return $"{playback?.Id}|{recording?.Id}|{playbackComm.Id}|{recordingComm.Id}|{bt}|{profileCount}|{drift}|{deviceFp}";
    }

    private void RegisterProfileHotkeys()
    {
        _hotkeyService?.UnregisterAll();
        foreach (var profile in ProfileService.GetAll())
        {
            if (profile.HotkeyKey == 0) continue;
            _hotkeyService?.Register(profile.HotkeyModifiers, profile.HotkeyKey, () =>
            {
                try
                {
                    if (!((App)Application.Current).TryUserApplyProfile(profile)) return;
                    var result = ProfileApplyService.Apply(profile);
                    ((App)Application.Current).MarkOwnChange();
                    LoadDevices();
                    LoadProfiles();
                    _miniWindow?.LoadProfiles();
                    ((App)Application.Current).NotifyProfileApplied(profile, result);
                }
                catch { }
            });
        }
    }

    private void ToggleMiniWindow_Click(object sender, RoutedEventArgs e)
    {
        if (_miniWindow is { IsLoaded: true })
        {
            _miniWindow.Close();
            _miniWindow = null;
            var s = SettingsService.Load();
            s.MiniWindowVisible = false;
            SettingsService.Save();
            UpdateMiniMenuCheck();
        }
        else
        {
            OpenMiniWindow();
        }
    }

    // ── Devices ──────────────────────────────────────────────

    private void LoadDevices()
    {
        // Drop stale references — buttons get rebuilt below and may not include the bar.
        _peakPlaybackTrack = null;
        _peakPlaybackFill = null;
        _peakRecordingTrack = null;
        _peakRecordingFill = null;
        _peakPlaybackDisplayed = 0;
        _peakRecordingDisplayed = 0;

        LoadPlaybackDevices();
        LoadRecordingDevices();
        RefreshPeakDevices();
        BluetoothWarning.Visibility = AudioDeviceService.HasBluetoothDevice()
            ? Visibility.Visible : Visibility.Collapsed;
        RefreshVoicemeeterStatus();
    }

    // Off-thread to keep the UI responsive: VBVMR_Login + GetParameterStringA serialized
    // through Voicemeeter's IPC takes ~100ms even when the app is local.
    private void RefreshVoicemeeterStatus()
    {
        Task.Run(() =>
        {
            VoicemeeterIoState? state = null;
            try { state = VoicemeeterService.GetIoState(); } catch { }
            if (state != null) state = EnforceVoicemeeterLock(state);
            Dispatcher.BeginInvoke(() => RenderVoicemeeterStatus(state));
        });
    }

    // If the Voicemeeter mute lock is on, push our snapshot back to the engine for any
    // strip whose mute state has drifted, then patch the returned state to reflect the
    // restored values so the UI doesn't briefly flash the drifted state.
    private static VoicemeeterIoState EnforceVoicemeeterLock(VoicemeeterIoState state)
    {
        var settings = SettingsService.Load();
        if (!settings.VoicemeeterMuteLocked) return state;
        var snap = settings.VoicemeeterStripMuteSnapshot;
        if (snap == null || snap.Count == 0) return state;

        bool anyDrift = false;
        var newInputs = new List<VoicemeeterIoSlot>(state.Inputs.Count);
        for (int i = 0; i < state.Inputs.Count; i++)
        {
            var slot = state.Inputs[i];
            if (i < snap.Count && slot.Muted != snap[i])
            {
                try { VoicemeeterService.SetStripMute(slot.Index, snap[i]); } catch { }
                anyDrift = true;
                newInputs.Add(slot with { Muted = snap[i] });
            }
            else
            {
                newInputs.Add(slot);
            }
        }
        return anyDrift ? state with { Inputs = newInputs } : state;
    }

    private void RenderVoicemeeterStatus(VoicemeeterIoState? state)
    {
        VoicemeeterIoList.Items.Clear();
        _vmOutputTracks = null; _vmOutputFills = null;
        _vmInputTracks = null; _vmInputFills = null;
        _vmOutputDisplayed = null; _vmInputDisplayed = null;
        _vmOutCount = 0; _vmInCount = 0;
        _lastVoicemeeterState = state;

        if (state == null)
        {
            VoicemeeterPanel.Visibility = Visibility.Collapsed;
            return;
        }
        VoicemeeterPanel.Visibility = Visibility.Visible;
        VoicemeeterTypeText.Text = state.TypeName;
        UpdateVoicemeeterLockButton();

        _vmOutCount = state.Outputs.Count;
        _vmInCount = state.Inputs.Count;
        _vmOutputTracks = new FrameworkElement[_vmOutCount];
        _vmOutputFills = new FrameworkElement[_vmOutCount];
        _vmInputTracks = new FrameworkElement[_vmInCount];
        _vmInputFills = new FrameworkElement[_vmInCount];
        _vmOutputDisplayed = new double[_vmOutCount];
        _vmInputDisplayed = new double[_vmInCount];

        for (int i = 0; i < state.Outputs.Count; i++)
        {
            var slot = state.Outputs[i];
            var (row, track, fill) = BuildVoicemeeterRow($"▶ {FormatSlotLabel(slot)}", slot.DeviceName, false, null, slot.DeviceMissing);
            VoicemeeterIoList.Items.Add(row);
            _vmOutputTracks[i] = track; _vmOutputFills[i] = fill;
        }
        for (int i = 0; i < state.Inputs.Count; i++)
        {
            var slot = state.Inputs[i];
            bool effectiveMuted = slot.Muted;
            if (_pendingMuteWrites.TryGetValue(slot.Index, out var pending))
            {
                if (DateTime.UtcNow - pending.When < PendingMuteWindow)
                    effectiveMuted = pending.Muted;
                else
                    _pendingMuteWrites.Remove(slot.Index);
            }
            var (row, track, fill) = BuildVoicemeeterRow($"● {FormatSlotLabel(slot)}", slot.DeviceName, effectiveMuted, slot.Index, slot.DeviceMissing);
            VoicemeeterIoList.Items.Add(row);
            _vmInputTracks[i] = track; _vmInputFills[i] = fill;
        }
    }

    private static string FormatSlotLabel(VoicemeeterIoSlot slot) =>
        string.IsNullOrEmpty(slot.CustomLabel) ? slot.Label : slot.CustomLabel;

    private (UIElement Row, FrameworkElement Track, FrameworkElement Fill) BuildVoicemeeterRow(
        string label, string? deviceName, bool muted, int? stripIndex, bool deviceMissing)
    {
        var stack = new StackPanel { Margin = new Thickness(0, 1, 0, 2) };

        var dock = new DockPanel();
        var labelText = new TextBlock
        {
            Text = label,
            FontSize = 11,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x5B, 0x21, 0xB6)),
            MinWidth = 64,
            Margin = new Thickness(0, 0, 8, 0),
            TextTrimming = TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
        };
        DockPanel.SetDock(labelText, Dock.Left);
        dock.Children.Add(labelText);

        var devText = new TextBlock
        {
            Text = string.IsNullOrEmpty(deviceName) ? "(未设置)" : deviceName,
            FontSize = 11,
            Foreground = new SolidColorBrush(string.IsNullOrEmpty(deviceName)
                ? Color.FromRgb(0x9C, 0xA3, 0xAF)
                : deviceMissing
                    ? Color.FromRgb(0xB9, 0x1C, 0x1C)
                    : Color.FromRgb(0x37, 0x41, 0x51)),
            TextTrimming = TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
        };

        if (stripIndex.HasValue)
        {
            var muteBadge = new Border
            {
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(5, 0, 5, 0),
                Margin = new Thickness(0, 0, 6, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Cursor = Cursors.Hand,
                Child = new TextBlock { Text = "静音", FontSize = 10 },
            };
            muteBadge.Tag = (stripIndex.Value, muted, devText);
            ApplyMuteVisual(muteBadge, devText, muted);
            muteBadge.MouseLeftButtonUp += MuteBadge_Click;
            DockPanel.SetDock(muteBadge, Dock.Left);
            dock.Children.Add(muteBadge);
        }

        if (deviceMissing)
        {
            var missingBadge = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0xFE, 0xE2, 0xE2)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xFC, 0xA5, 0xA5)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(5, 0, 5, 0),
                Margin = new Thickness(0, 0, 6, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = "已丢失",
                    FontSize = 10,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(Color.FromRgb(0xB9, 0x1C, 0x1C)),
                },
                ToolTip = "Voicemeeter 配置的设备在系统中找不到",
            };
            DockPanel.SetDock(missingBadge, Dock.Left);
            dock.Children.Add(missingBadge);
        }

        dock.Children.Add(devText);
        stack.Children.Add(dock);

        var fill = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(0x7C, 0x3A, 0xED)),
            CornerRadius = new CornerRadius(1.5),
            HorizontalAlignment = HorizontalAlignment.Left,
            Width = 0,
        };
        var track = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(0xDD, 0xD6, 0xFE)),
            CornerRadius = new CornerRadius(1.5),
            Height = 3,
            Margin = new Thickness(64, 2, 2, 0),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Child = fill,
            ClipToBounds = true,
        };
        stack.Children.Add(track);

        return (stack, track, fill);
    }

    private void MuteBadge_Click(object sender, MouseButtonEventArgs e)
    {
        if (sender is not Border b) return;
        if (b.Tag is not ValueTuple<int, bool, TextBlock> tag) return;
        var (idx, wasMuted, devText) = tag;
        bool newMuted = !wasMuted;
        if (!VoicemeeterService.SetStripMute(idx, newMuted)) return;
        // Optimistic update + record latest intent so any in-flight dirty poll that
        // reads stale state won't visually revert us.
        _pendingMuteWrites[idx] = (newMuted, DateTime.UtcNow);
        ApplyMuteVisual(b, devText, newMuted);
        b.Tag = (idx, newMuted, devText);
        // If the Voicemeeter is locked, our intent becomes the new snapshot value —
        // otherwise the next poll would revert the change we just made via the badge.
        var settings = SettingsService.Load();
        if (settings.VoicemeeterMuteLocked && idx >= 0 && idx < settings.VoicemeeterStripMuteSnapshot.Count)
        {
            settings.VoicemeeterStripMuteSnapshot[idx] = newMuted;
            SettingsService.Save();
        }
    }

    private void UpdateVoicemeeterLockButton()
    {
        bool locked = SettingsService.Load().VoicemeeterMuteLocked;
        VoicemeeterLockBtn.Content = locked ? "\U0001F512" : "\U0001F513";
        VoicemeeterLockBtn.Foreground = locked
            ? new SolidColorBrush(Color.FromRgb(0xD9, 0x77, 0x06))
            : new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF));
        VoicemeeterLockBtn.ToolTip = locked
            ? "Strip 静音已锁定 — 点击解锁"
            : "锁定当前 Strip 静音状态（外部修改会自动恢复）";
    }

    private void VoicemeeterLock_Click(object sender, RoutedEventArgs e)
    {
        var settings = SettingsService.Load();
        if (settings.VoicemeeterMuteLocked)
        {
            settings.VoicemeeterMuteLocked = false;
            settings.VoicemeeterStripMuteSnapshot = [];
        }
        else
        {
            // Snapshot from the most recent rendered state — fallback to a fresh fetch.
            var state = _lastVoicemeeterState;
            if (state == null)
            {
                try { state = VoicemeeterService.GetIoState(); } catch { }
            }
            if (state == null) return;
            settings.VoicemeeterMuteLocked = true;
            settings.VoicemeeterStripMuteSnapshot = state.Inputs.Select(s => s.Muted).ToList();
        }
        SettingsService.Save();
        UpdateVoicemeeterLockButton();
    }

    private static void ApplyMuteVisual(Border badge, TextBlock devText, bool muted)
    {
        badge.Background = muted
            ? (Brush)new SolidColorBrush(Color.FromRgb(0xFE, 0xE2, 0xE2))
            : Brushes.Transparent;
        badge.BorderBrush = muted
            ? (Brush)new SolidColorBrush(Color.FromRgb(0xFC, 0xA5, 0xA5))
            : (Brush)new SolidColorBrush(Color.FromRgb(0xC4, 0xB5, 0xFD));
        badge.ToolTip = muted ? "点击取消静音" : "点击静音";
        if (badge.Child is TextBlock txt)
        {
            txt.FontWeight = muted ? FontWeights.SemiBold : FontWeights.Normal;
            txt.Foreground = muted
                ? (Brush)new SolidColorBrush(Color.FromRgb(0xB9, 0x1C, 0x1C))
                : (Brush)new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF));
        }
        if (muted)
        {
            devText.TextDecorations = TextDecorations.Strikethrough;
            devText.Opacity = 0.6;
        }
        else
        {
            devText.TextDecorations = null;
            devText.Opacity = 1.0;
        }
    }

    private void LoadPlaybackDevices()
    {
        PlaybackList.Items.Clear();
        var settings = SettingsService.Load();
        var hidden = new HashSet<string>(settings.HiddenDeviceIds, StringComparer.OrdinalIgnoreCase);
        int hiddenCount = 0;
        foreach (var device in AudioDeviceService.GetPlaybackDevices(settings.ShowDisabledDevices))
        {
            bool isHidden = hidden.Contains(device.Id);
            if (isHidden) hiddenCount++;
            if (isHidden && !settings.ShowHiddenDevices) continue;
            bool isSelected = device.Id == _selectedPlaybackId;
            var btn = CreateDeviceButton(device, isHidden, isSelected, isPlayback: true);
            PlaybackList.Items.Add(btn);
        }
        UpdateHiddenCount(PlaybackHiddenCount, hiddenCount);
    }

    private void LoadRecordingDevices()
    {
        RecordingList.Items.Clear();
        var settings = SettingsService.Load();
        var hidden = new HashSet<string>(settings.HiddenDeviceIds, StringComparer.OrdinalIgnoreCase);
        int hiddenCount = 0;
        foreach (var device in AudioDeviceService.GetRecordingDevices(settings.ShowDisabledDevices))
        {
            bool isHidden = hidden.Contains(device.Id);
            if (isHidden) hiddenCount++;
            if (isHidden && !settings.ShowHiddenDevices) continue;
            bool isSelected = device.Id == _selectedRecordingId;
            var btn = CreateDeviceButton(device, isHidden, isSelected, isPlayback: false);
            RecordingList.Items.Add(btn);
        }
        UpdateHiddenCount(RecordingHiddenCount, hiddenCount);
    }

    private static void UpdateHiddenCount(TextBlock block, int count)
    {
        if (count > 0)
        {
            block.Text = $"({count} \u5DF2\u9690\u85CF)";
            block.Visibility = Visibility.Visible;
        }
        else
        {
            block.Visibility = Visibility.Collapsed;
        }
    }

    private Button CreateDeviceButton(AudioDeviceInfo device, bool isHidden = false, bool isSelected = false, bool isPlayback = true)
    {
        // Blue highlight = user selection only. Default device uses plain style + checkmark.
        var style = isSelected
            ? (Style)FindResource("DefaultDeviceButton")
            : (Style)FindResource("DeviceButton");

        var settings = SettingsService.Load();
        settings.DeviceNicknames.TryGetValue(device.Id, out var nickname);
        string displayName = !string.IsNullOrEmpty(nickname) ? nickname : device.Name;

        // Root = DockPanel. Indicator (✓/●) docked to the right so it doesn't shift text.
        var root = new DockPanel { LastChildFill = true };

        var indicator = new TextBlock
        {
            FontWeight = FontWeights.SemiBold,
            FontSize = 14,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
            Width = 20,
            Margin = new Thickness(6, 0, 2, 0),
            TextAlignment = TextAlignment.Center,
        };
        if (device.IsDefault && !device.IsDisabled)
        {
            indicator.Text = "\u2714";
            indicator.Foreground = new SolidColorBrush(Color.FromRgb(0x10, 0x9E, 0x7A));
        }
        else if (isSelected && !device.IsDisabled)
        {
            indicator.Text = "\u25CF";
            indicator.Foreground = new SolidColorBrush(Color.FromRgb(0x3B, 0x82, 0xF6));
        }
        DockPanel.SetDock(indicator, Dock.Right);
        root.Children.Add(indicator);

        var panel = new StackPanel { Orientation = Orientation.Horizontal };

        if (device.Icon != null)
        {
            var img = new Image
            {
                Source = device.Icon,
                Width = 48,
                Height = 48,
                Margin = new Thickness(0, 0, 6, 0),
            };
            RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.HighQuality);
            panel.Children.Add(img);
        }

        var nameStack = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
        var nameRow = new StackPanel { Orientation = Orientation.Horizontal };
        nameRow.Children.Add(new TextBlock { Text = displayName, VerticalAlignment = VerticalAlignment.Center });
        if (device.IsDisabled)
        {
            nameRow.Children.Add(new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0xF3, 0xF4, 0xF6)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xE5, 0xE7, 0xEB)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(5, 1, 5, 1),
                Margin = new Thickness(8, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = "\u5DF2\u7981\u7528",
                    FontSize = 10,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80)),
                },
            });
        }
        nameStack.Children.Add(nameRow);
        if (!string.IsNullOrEmpty(nickname))
        {
            nameStack.Children.Add(new TextBlock
            {
                Text = device.Name,
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80)),
            });
        }
        panel.Children.Add(nameStack);
        root.Children.Add(panel);

        double opacity = 1.0;
        if (device.IsDisabled) opacity = 0.55;
        else if (isHidden) opacity = 0.5;

        object content = root;
        if (device.IsDefault && !device.IsDisabled)
        {
            var stack = new StackPanel();
            stack.Children.Add(root);

            var fill = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0x10, 0x9E, 0x7A)),
                CornerRadius = new CornerRadius(2),
                HorizontalAlignment = HorizontalAlignment.Left,
                Width = 0,
            };
            var track = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0xE5, 0xE7, 0xEB)),
                CornerRadius = new CornerRadius(2),
                Height = 4,
                Margin = new Thickness(2, 6, 2, 0),
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Child = fill,
                ClipToBounds = true,
            };
            stack.Children.Add(track);

            if (isPlayback) { _peakPlaybackTrack = track; _peakPlaybackFill = fill; }
            else { _peakRecordingTrack = track; _peakRecordingFill = fill; }

            content = stack;
        }

        var btn = new Button
        {
            Content = content,
            Tag = device.Id,
            Style = style,
            Opacity = opacity,
            Cursor = device.IsDisabled ? System.Windows.Input.Cursors.Arrow : System.Windows.Input.Cursors.Hand,
        };
        btn.Click += DeviceButton_Click;

        bool anyLocked = ((App)Application.Current).IsAnyProfileLocked();

        var menu = new ContextMenu();
        var setDefaultItem = new MenuItem { Header = "设为默认设备", Tag = device.Id };
        setDefaultItem.Click += SetDefaultDevice_Click;
        setDefaultItem.IsEnabled = !device.IsDefault && !device.IsDisabled && !anyLocked;
        if (anyLocked) setDefaultItem.ToolTip = "已锁定 — 请先解锁配置";
        menu.Items.Add(setDefaultItem);
        var toggleEnableItem = new MenuItem
        {
            Header = device.IsDisabled ? "启用此设备" : "禁用此设备",
            Tag = device.Id,
            IsEnabled = !anyLocked,
            ToolTip = anyLocked ? "已锁定 — 请先解锁配置" : null,
        };
        toggleEnableItem.Click += ToggleEnableDevice_Click;
        menu.Items.Add(toggleEnableItem);
        menu.Items.Add(new Separator());
        var renameItem = new MenuItem { Header = "重命名…", Tag = device };
        renameItem.Click += RenameDevice_Click;
        menu.Items.Add(renameItem);
        var toggleItem = new MenuItem { Header = isHidden ? "显示此设备" : "隐藏此设备", Tag = device.Id };
        toggleItem.Click += ToggleHideDevice_Click;
        menu.Items.Add(toggleItem);
        btn.ContextMenu = menu;

        return btn;
    }

    private void RenameDevice_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem mi || mi.Tag is not AudioDeviceInfo device) return;
        var settings = SettingsService.Load();
        settings.DeviceNicknames.TryGetValue(device.Id, out var current);

        // Dedup scope = same flow only (playback vs recording can share names)
        var playback = AudioDeviceService.GetPlaybackDevices();
        var sameCategory = playback.Any(d => string.Equals(d.Id, device.Id, StringComparison.OrdinalIgnoreCase))
            ? playback
            : AudioDeviceService.GetRecordingDevices();

        string? proposed = current;
        while (true)
        {
            var dlg = new DeviceRenameDialog(device.Name, proposed) { Owner = this };
            if (dlg.ShowDialog() != true) return;

            if (string.IsNullOrEmpty(dlg.Nickname))
            {
                settings.DeviceNicknames.Remove(device.Id);
                SettingsService.Save();
                LoadDevices();
                return;
            }

            bool conflict = sameCategory.Any(d =>
            {
                if (string.Equals(d.Id, device.Id, StringComparison.OrdinalIgnoreCase)) return false;
                var displayed = settings.DeviceNicknames.TryGetValue(d.Id, out var nn) && !string.IsNullOrEmpty(nn)
                    ? nn : d.Name;
                return string.Equals(displayed, dlg.Nickname, StringComparison.OrdinalIgnoreCase);
            });

            if (conflict)
            {
                MessageBox.Show($"\u540C\u7C7B\u8BBE\u5907\u4E2D\u5DF2\u5B58\u5728\u540D\u79F0\u300C{dlg.Nickname}\u300D\uFF0C\u8BF7\u6362\u4E00\u4E2A\u540D\u79F0\u3002",
                    "\u63D0\u793A", MessageBoxButton.OK, MessageBoxImage.Warning);
                proposed = dlg.Nickname;
                continue;
            }

            settings.DeviceNicknames[device.Id] = dlg.Nickname;
            SettingsService.Save();
            LoadDevices();
            return;
        }
    }

    private void ToggleHideDevice_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem mi || mi.Tag is not string deviceId) return;
        var settings = SettingsService.Load();
        var hidden = new HashSet<string>(settings.HiddenDeviceIds, StringComparer.OrdinalIgnoreCase);
        if (hidden.Contains(deviceId)) hidden.Remove(deviceId);
        else hidden.Add(deviceId);
        settings.HiddenDeviceIds = hidden.ToList();
        SettingsService.Save();
        LoadDevices();
    }

    private void ShowHidden_Changed(object sender, RoutedEventArgs e)
    {
        if (!IsLoaded) return;
        var settings = SettingsService.Load();
        settings.ShowHiddenDevices = ShowHiddenCheck.IsChecked == true;
        SettingsService.Save();
        LoadDevices();
    }

    private void ShowDisabled_Changed(object sender, RoutedEventArgs e)
    {
        if (!IsLoaded) return;
        var settings = SettingsService.Load();
        settings.ShowDisabledDevices = ShowDisabledCheck.IsChecked == true;
        SettingsService.Save();
        LoadDevices();
    }

    private void DeviceButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not string deviceId) return;

        // Left-click = toggle selection (for profile creation). Right-click opens context menu for "set default".
        var showDisabled = SettingsService.Load().ShowDisabledDevices;
        var playback = AudioDeviceService.GetPlaybackDevices(showDisabled);
        var recording = AudioDeviceService.GetRecordingDevices(showDisabled);
        var dev = playback.Find(d => d.Id == deviceId) ?? recording.Find(d => d.Id == deviceId);
        if (dev == null || dev.IsDisabled) return;

        bool isPlayback = playback.Any(d => d.Id == deviceId);
        if (isPlayback)
            _selectedPlaybackId = _selectedPlaybackId == deviceId ? null : deviceId;
        else
            _selectedRecordingId = _selectedRecordingId == deviceId ? null : deviceId;

        LoadDevices();
    }

    private void SetDefaultDevice_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem mi || mi.Tag is not string deviceId) return;
        if (!((App)Application.Current).TryUserChangeDevice()) return;
        try
        {
            AudioDeviceService.SetDefaultDevice(deviceId);
            ((App)Application.Current).MarkOwnChange();
            LoadDevices();
            LoadProfiles();
            _miniWindow?.LoadProfiles();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"\u5207\u6362\u8BBE\u5907\u5931\u8D25\uFF1A\n{ex.Message}", "\u9519\u8BEF",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ToggleEnableDevice_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem mi || mi.Tag is not string deviceId) return;
        if (!((App)Application.Current).TryUserChangeDevice()) return;
        var dev = AudioDeviceService.GetPlaybackDevices(true).Find(d => d.Id == deviceId)
               ?? AudioDeviceService.GetRecordingDevices(true).Find(d => d.Id == deviceId);
        if (dev == null) return;

        try
        {
            AudioDeviceService.SetDeviceEnabled(deviceId, dev.IsDisabled);
            ((App)Application.Current).MarkOwnChange();
            LoadDevices();
            LoadProfiles();
            _miniWindow?.LoadProfiles();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"\u64CD\u4F5C\u5931\u8D25\uFF1A\n{ex.Message}", "\u9519\u8BEF",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ── Profiles ─────────────────────────────────────────────

    private void LoadProfiles()
    {
        ProfileList.Items.Clear();

        var profiles = ProfileService.GetAll();
        ProfileEmptyHint.Visibility = profiles.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        var allPlayback = AudioDeviceService.GetPlaybackDevices();
        var allRecording = AudioDeviceService.GetRecordingDevices();
        var currentPlayback = allPlayback.Find(d => d.IsDefault);
        var currentRecording = allRecording.Find(d => d.IsDefault);
        var currentPlaybackComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Render);
        var currentRecordingComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Capture);

        var nicknames = SettingsService.Load().DeviceNicknames;
        string ResolveDeviceName(string? id, string? rawName)
        {
            var raw = rawName ?? "\u65E0";
            if (!string.IsNullOrEmpty(id)
                && nicknames.TryGetValue(id, out var nn)
                && !string.IsNullOrEmpty(nn))
                return $"{nn} ({raw})";
            return raw;
        }

        bool anyActive = false;

        foreach (var profile in profiles)
        {
            bool isActive = profile.PlaybackDeviceId == currentPlayback?.Id
                         && profile.RecordingDeviceId == currentRecording?.Id
                         && profile.PlaybackDeviceId == currentPlaybackComm.Id
                         && profile.RecordingDeviceId == currentRecordingComm.Id;
            if (isActive) anyActive = true;

            var lockedSettingsId = SettingsService.Load().LockedProfileId;
            bool isLocked = lockedSettingsId == profile.Id;
            bool blockedByLock = lockedSettingsId.HasValue && !isLocked;

            // Left side: name + device info
            var nameRow = new StackPanel { Orientation = Orientation.Horizontal };
            if (!string.IsNullOrEmpty(profile.Color))
            {
                try
                {
                    var color = (Color)ColorConverter.ConvertFromString(profile.Color);
                    nameRow.Children.Add(new System.Windows.Shapes.Ellipse
                    {
                        Width = 10,
                        Height = 10,
                        Fill = new SolidColorBrush(color),
                        Margin = new Thickness(0, 0, 6, 0),
                        VerticalAlignment = VerticalAlignment.Center,
                    });
                }
                catch { }
            }
            if (isLocked)
            {
                nameRow.Children.Add(new TextBlock
                {
                    Text = "🔒 ",
                    FontSize = 12,
                    Foreground = new SolidColorBrush(Color.FromRgb(0xD9, 0x77, 0x06)),
                    VerticalAlignment = VerticalAlignment.Center,
                });
            }
            var nameText = new TextBlock
            {
                Text = (isActive ? "\u2714 " : "") + profile.Name,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = isActive
                    ? new SolidColorBrush(Color.FromRgb(0x1E, 0x40, 0xAF))
                    : new SolidColorBrush(Color.FromRgb(0x11, 0x18, 0x27)),
            };
            nameRow.Children.Add(nameText);

            bool playbackMissing = !string.IsNullOrEmpty(profile.PlaybackDeviceId)
                && !allPlayback.Any(d => string.Equals(d.Id, profile.PlaybackDeviceId, StringComparison.Ordinal));
            bool recordingMissing = !string.IsNullOrEmpty(profile.RecordingDeviceId)
                && !allRecording.Any(d => string.Equals(d.Id, profile.RecordingDeviceId, StringComparison.Ordinal));
            var missingBrush = new SolidColorBrush(Color.FromRgb(0xD3, 0x2F, 0x2F));

            var detailText = new TextBlock
            {
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80)),
                Margin = new Thickness(0, 3, 0, 0),
                TextTrimming = TextTrimming.CharacterEllipsis,
            };

            detailText.Inlines.Add(new Run($"▶ {ResolveDeviceName(profile.PlaybackDeviceId, profile.PlaybackDeviceName)}"));
            if (playbackMissing)
                detailText.Inlines.Add(new Run(" 未连接") { Foreground = missingBrush, FontWeight = FontWeights.SemiBold });
            detailText.Inlines.Add(new Run($"  ·  ● {ResolveDeviceName(profile.RecordingDeviceId, profile.RecordingDeviceName)}"));
            if (recordingMissing)
                detailText.Inlines.Add(new Run(" 未连接") { Foreground = missingBrush, FontWeight = FontWeights.SemiBold });
            if (profile.HotkeyKey != 0)
                detailText.Inlines.Add(new Run($"  ·  ⌨ {FormatHotkey(profile.HotkeyModifiers, profile.HotkeyKey)}"));

            var infoPanel = new StackPanel();
            infoPanel.Children.Add(nameRow);
            infoPanel.Children.Add(detailText);

            var applyBtn = new Button
            {
                Content = infoPanel,
                Style = (Style)FindResource("ProfileApplyButton"),
                Tag = profile.Id,
            };
            applyBtn.Click += ApplyProfile_Click;

            // Right side: lock + edit + delete links
            var lockBtn = new Button
            {
                Content = isLocked ? "\U0001F512" : "\U0001F513",
                Style = (Style)FindResource("SmallLinkButton"),
                Tag = profile.Id,
                Foreground = isLocked
                    ? new SolidColorBrush(Color.FromRgb(0xD9, 0x77, 0x06))
                    : new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF)),
                ToolTip = isLocked ? "已锁定 — 点击解锁" : "锁定到此配置",
                FontSize = 13,
            };
            lockBtn.Click += LockProfile_Click;

            var editBtn = new Button
            {
                Content = "\u7F16\u8F91",
                Style = (Style)FindResource("SmallLinkButton"),
                Tag = profile.Id,
            };
            editBtn.Click += EditProfile_Click;

            var separator = new TextBlock
            {
                Text = "|",
                Foreground = new SolidColorBrush(Color.FromRgb(0xD1, 0xD5, 0xDB)),
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 11,
            };

            var deleteBtn = new Button
            {
                Content = "\u5220\u9664",
                Style = (Style)FindResource("SmallLinkButton"),
                Tag = profile.Id,
                Foreground = new SolidColorBrush(Color.FromRgb(0xDC, 0x26, 0x26)),
            };
            deleteBtn.Click += DeleteProfile_Click;

            var actionPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center,
            };
            actionPanel.Children.Add(lockBtn);
            actionPanel.Children.Add(new TextBlock
            {
                Text = "|",
                Foreground = new SolidColorBrush(Color.FromRgb(0xD1, 0xD5, 0xDB)),
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 11,
                Margin = new Thickness(2, 0, 2, 0),
            });
            actionPanel.Children.Add(editBtn);
            actionPanel.Children.Add(separator);
            actionPanel.Children.Add(deleteBtn);

            // Drag handle (only this region initiates a drag; rest of card stays clickable to apply)
            var grip = new TextBlock
            {
                Text = "☰",
                FontSize = 14,
                Foreground = new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF)),
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Cursor = Cursors.SizeAll,
                Tag = profile.Id,
                ToolTip = "拖拽调整顺序",
            };
            grip.PreviewMouseLeftButtonDown += ProfileGrip_MouseDown;
            grip.PreviewMouseMove += ProfileGrip_MouseMove;

            // Card layout
            var dock = new DockPanel();
            DockPanel.SetDock(grip, Dock.Left);
            dock.Children.Add(grip);
            DockPanel.SetDock(actionPanel, Dock.Right);
            dock.Children.Add(actionPanel);
            dock.Children.Add(applyBtn);

            var card = new Border
            {
                Style = (Style)FindResource(isActive ? "ActiveProfileCard" : "ProfileCard"),
                Child = dock,
                Cursor = blockedByLock ? System.Windows.Input.Cursors.Arrow : System.Windows.Input.Cursors.Hand,
                Tag = profile.Id,
                AllowDrop = true,
                Opacity = blockedByLock ? 0.5 : 1.0,
                ToolTip = blockedByLock ? "已被锁定到其他配置 — 先解锁才能切换" : null,
            };
            card.MouseLeftButtonUp += (_, e) =>
            {
                if (e.ChangedButton != System.Windows.Input.MouseButton.Left) return;
                ApplyProfile_Click(applyBtn, e);
            };
            card.PreviewDragOver += ProfileCard_DragOver;
            card.Drop += ProfileCard_Drop;

            ProfileList.Items.Add(card);
        }

        bool showWarning = profiles.Count > 0 && !anyActive;
        ProfileWarningText.Visibility = showWarning ? Visibility.Visible : Visibility.Collapsed;
        ProfileWarningText.Tag = (showWarning && SettingsService.Load().EnableBlinkAnimation) ? "Blink" : null;

        var drifted = ((App)Application.Current).DriftedApps;
        if (drifted.Count > 0)
        {
            AppDriftWarningText.Text = $"{drifted.Count} \u4E2A\u5E94\u7528\u504F\u79BB: {string.Join(", ", drifted)}";
            AppDriftWarningText.Visibility = Visibility.Visible;
        }
        else
        {
            AppDriftWarningText.Visibility = Visibility.Collapsed;
        }
    }

    private void ProfileGrip_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not TextBlock tb || tb.Tag is not Guid id) return;
        _profileDragStart = e.GetPosition(this);
        _profileDragSourceId = id;
    }

    private void ProfileGrip_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed) return;
        if (_profileDragSourceId == Guid.Empty) return;
        if (sender is not TextBlock tb) return;

        var pos = e.GetPosition(this);
        if (Math.Abs(pos.X - _profileDragStart.X) < SystemParameters.MinimumHorizontalDragDistance
         && Math.Abs(pos.Y - _profileDragStart.Y) < SystemParameters.MinimumVerticalDragDistance) return;

        try
        {
            DragDrop.DoDragDrop(tb, new DataObject("ProfileId", _profileDragSourceId), DragDropEffects.Move);
        }
        finally
        {
            _profileDragSourceId = Guid.Empty;
        }
    }

    private void ProfileCard_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent("ProfileId") ? DragDropEffects.Move : DragDropEffects.None;
        e.Handled = true;
    }

    private void ProfileCard_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent("ProfileId")) return;
        if (e.Data.GetData("ProfileId") is not Guid sourceId) return;
        if (sender is not Border target || target.Tag is not Guid targetId) return;
        if (sourceId == targetId) return;

        var profiles = ProfileService.GetAll();
        var src = profiles.Find(p => p.Id == sourceId);
        if (src == null) return;
        profiles.Remove(src);

        int targetIdx = profiles.FindIndex(p => p.Id == targetId);
        if (targetIdx < 0) { profiles.Add(src); }
        else
        {
            // Drop above or below depending on cursor position within the target card
            var pos = e.GetPosition(target);
            if (pos.Y > target.ActualHeight / 2) targetIdx++;
            profiles.Insert(targetIdx, src);
        }

        for (int i = 0; i < profiles.Count; i++) profiles[i].Order = i + 1;
        ProfileService.SaveAll(profiles);

        LoadProfiles();
        _miniWindow?.LoadProfiles();
    }

    private void ApplyProfile_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not Guid id) return;

        var profile = ProfileService.GetAll().Find(p => p.Id == id);
        if (profile == null) return;

        if (!((App)Application.Current).TryUserApplyProfile(profile)) return;

        try
        {
            var result = ProfileApplyService.Apply(profile);
            ((App)Application.Current).MarkOwnChange();
            LoadDevices();
            LoadProfiles();
            _miniWindow?.LoadProfiles();
            ((App)Application.Current).NotifyProfileApplied(profile, result);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"\u5E94\u7528\u914D\u7F6E\u5931\u8D25\uFF1A\n{ex.Message}", "\u9519\u8BEF",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void NewProfile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new ProfileEditDialog { Owner = this };
        if (dialog.ShowDialog() != true) return;

        // Use user-selected devices if set, otherwise fall back to current defaults.
        var allPlayback = AudioDeviceService.GetPlaybackDevices();
        var allRecording = AudioDeviceService.GetRecordingDevices();
        var playback = _selectedPlaybackId != null
            ? allPlayback.Find(d => d.Id == _selectedPlaybackId)
            : allPlayback.Find(d => d.IsDefault);
        var recording = _selectedRecordingId != null
            ? allRecording.Find(d => d.Id == _selectedRecordingId)
            : allRecording.Find(d => d.IsDefault);

        var profile = new DeviceProfile
        {
            Name = dialog.ProfileName,
            PlaybackDeviceId = playback?.Id,
            PlaybackDeviceName = playback?.Name,
            RecordingDeviceId = recording?.Id,
            RecordingDeviceName = recording?.Name,
            HotkeyModifiers = dialog.HotkeyModifiers,
            HotkeyKey = dialog.HotkeyKey,
            AppOverrides = dialog.AppOverrides,
            RestartVoicemeeterAfterApply = dialog.RestartVoicemeeter,
            ShowInMiniWindow = dialog.ShowInMiniWindow,
            Color = dialog.SelectedColor,
        };
        ProfileService.Save(profile);
        RegisterProfileHotkeys();

        _selectedPlaybackId = null;
        _selectedRecordingId = null;

        LoadDevices();
        LoadProfiles();
        _miniWindow?.LoadProfiles();
    }

    private void EditProfile_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not Guid id) return;

        var profile = ProfileService.GetAll().Find(p => p.Id == id);
        if (profile == null) return;

        var dialog = new ProfileEditDialog(profile.Name, profile.HotkeyModifiers, profile.HotkeyKey, profile.AppOverrides, profile.RestartVoicemeeterAfterApply, profile.ShowInMiniWindow, profile.Color) { Owner = this };
        if (dialog.ShowDialog() != true) return;

        profile.Name = dialog.ProfileName;
        profile.HotkeyModifiers = dialog.HotkeyModifiers;
        profile.HotkeyKey = dialog.HotkeyKey;
        profile.AppOverrides = dialog.AppOverrides;
        profile.RestartVoicemeeterAfterApply = dialog.RestartVoicemeeter;
        profile.ShowInMiniWindow = dialog.ShowInMiniWindow;
        profile.Color = dialog.SelectedColor;
        ProfileService.Save(profile);
        RegisterProfileHotkeys();
        LoadProfiles();
        _miniWindow?.LoadProfiles();
    }

    private void DeleteProfile_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not Guid id) return;

        var profile = ProfileService.GetAll().Find(p => p.Id == id);
        if (profile == null) return;

        var result = MessageBox.Show($"\u786E\u5B9A\u5220\u9664\u914D\u7F6E \"{profile.Name}\"\uFF1F", "\u786E\u8BA4",
            MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result == MessageBoxResult.Yes)
        {
            ProfileService.Delete(id);
            // Clear lock if the deleted profile was the locked one.
            var s = SettingsService.Load();
            if (s.LockedProfileId == id)
            {
                s.LockedProfileId = null;
                SettingsService.Save();
            }
            RegisterProfileHotkeys();
            LoadProfiles();
            _miniWindow?.LoadProfiles();
        }
    }

    private void LockProfile_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not Guid id) return;
        var settings = SettingsService.Load();
        bool wasLocked = settings.LockedProfileId == id;
        settings.LockedProfileId = wasLocked ? null : id;
        SettingsService.Save();

        if (!wasLocked)
        {
            // Locking implies "make the system match this profile right now".
            var profile = ProfileService.GetAll().Find(p => p.Id == id);
            if (profile != null)
            {
                try
                {
                    ProfileApplyService.Apply(profile);
                    ((App)Application.Current).MarkOwnChange();
                }
                catch { }
            }
        }

        LoadDevices();
        LoadProfiles();
        _miniWindow?.LoadProfiles();
    }

    // ── Other ────────────────────────────────────────────────

    private void Refresh_Click(object sender, RoutedEventArgs e)
    {
        LoadProfiles();
        LoadDevices();
    }

    private void OpenSoundSettings_Click(object sender, RoutedEventArgs e)
    {
        Process.Start(new ProcessStartInfo("rundll32.exe", "shell32.dll,Control_RunDLL mmsys.cpl,,0")
        {
            UseShellExecute = false
        });
    }

    private void OpenSettings_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SettingsWindow { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _lastStateSignature = "";
            RefreshFromExternalChange();
            _miniWindow?.LoadProfiles();
        }
    }

    private void OpenAbout_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new AboutWindow { Owner = this };
        dlg.ShowDialog();
    }

    private void OpenAppAudio_Click(object sender, RoutedEventArgs e)
    {
        var win = new AppAudioWindow { Owner = this };
        win.ShowDialog();
    }

    private void OpenAppProfiles_Click(object sender, RoutedEventArgs e)
    {
        var win = new AppProfilesWindow { Owner = this };
        win.ShowDialog();
    }

    private void ExitApp_Click(object sender, RoutedEventArgs e)
    {
        var result = MessageBox.Show("\u786E\u5B9A\u9000\u51FA\u97F3\u9891\u5207\u6362\u52A9\u624B\uFF1F", "\u786E\u8BA4",
            MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result == MessageBoxResult.Yes)
        {
            ((App)Application.Current).ExitFromUI();
        }
    }

    private static string FormatHotkey(int modifiers, int key)
    {
        var parts = new List<string>();
        if ((modifiers & 0x0002) != 0) parts.Add("Ctrl");
        if ((modifiers & 0x0001) != 0) parts.Add("Alt");
        if ((modifiers & 0x0004) != 0) parts.Add("Shift");
        parts.Add(KeyInterop.KeyFromVirtualKey(key).ToString());
        return string.Join("+", parts);
    }
}
