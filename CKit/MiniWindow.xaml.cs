using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace AudioDeviceSwitcher;

public partial class MiniWindow : Window
{
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

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);
    private const int VK_LBUTTON = 0x01;
    private static bool IsLeftMouseDown() => (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;

    private readonly Action? _onProfileApplied;
    private readonly Border _border;
    private bool _isLocked;

    public MiniWindow(Action? onProfileApplied = null)
    {
        InitializeComponent();
        _onProfileApplied = onProfileApplied;
        _border = (Border)Content;

        _isLocked = LockCheck.IsChecked == true;
        ProfileList.Opacity = _isLocked ? 0.5 : 1.0;

        ApplySettings();
        LoadProfiles();

        // External device/profile/drift refreshes are pushed by App.CheckDeviceChanges
        // (event-driven via DeviceChangeNotifier + 5-sec safety-net timer). No local poll.

        Closed += (_, _) => SaveSettings();
        LocationChanged += (_, _) =>
        {
            if (IsLeftMouseDown()) SaveSettings();
        };
        SourceInitialized += (_, _) =>
        {
            var s = SettingsService.Load();
            if (s.MiniWindowLeft is double l && s.MiniWindowTop is double t)
            {
                var hwnd = new WindowInteropHelper(this).Handle;
                if (hwnd != IntPtr.Zero)
                    SetWindowPos(hwnd, IntPtr.Zero, (int)l, (int)t, 0, 0,
                        SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
            }
        };
    }

    private void ApplySettings()
    {
        // Position restored via Win32 in SourceInitialized (physical pixels).
        WindowStartupLocation = WindowStartupLocation.Manual;
        var s = SettingsService.Load();
        Opacity = s.MiniWindowOpacity;
        OpacitySlider.Value = s.MiniWindowOpacity;
    }

    private void SaveSettings()
    {
        var s = SettingsService.Load();
        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out var rect))
        {
            s.MiniWindowLeft = rect.Left;
            s.MiniWindowTop = rect.Top;
        }
        if (!double.IsNaN(Opacity) && !double.IsInfinity(Opacity)) s.MiniWindowOpacity = Opacity;
        SettingsService.Save();
    }

    public void LoadProfiles()
    {
        ProfileList.Items.Clear();

        var allProfiles = ProfileService.GetAll();
        var profiles = allProfiles.Where(p => p.ShowInMiniWindow).ToList();
        EmptyHint.Visibility = profiles.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

        var allPlayback = AudioDeviceService.GetPlaybackDevices();
        var allRecording = AudioDeviceService.GetRecordingDevices();
        var currentPlayback = allPlayback.Find(d => d.IsDefault);
        var currentRecording = allRecording.Find(d => d.IsDefault);
        var currentPlaybackComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Render);
        var currentRecordingComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Capture);

        var nicknames = SettingsService.Load().DeviceNicknames;
        string? NicknameOf(string? id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            return nicknames.TryGetValue(id, out var nn) && !string.IsNullOrEmpty(nn) ? nn : null;
        }

        foreach (var profile in profiles)
        {
            bool isActive = profile.PlaybackDeviceId == currentPlayback?.Id
                         && profile.RecordingDeviceId == currentRecording?.Id
                         && profile.PlaybackDeviceId == currentPlaybackComm.Id
                         && profile.RecordingDeviceId == currentRecordingComm.Id;

            var playNn = NicknameOf(profile.PlaybackDeviceId);
            var recNn = NicknameOf(profile.RecordingDeviceId);
            var nicknameParts = new List<string>();
            if (playNn != null) nicknameParts.Add($"\u25B6 {playNn}");
            if (recNn != null) nicknameParts.Add($"\u25CF {recNn}");
            string? nicknameLine = nicknameParts.Count > 0 ? string.Join("  ", nicknameParts) : null;

            bool playbackMissing = !string.IsNullOrEmpty(profile.PlaybackDeviceId)
                && !allPlayback.Any(d => string.Equals(d.Id, profile.PlaybackDeviceId, StringComparison.Ordinal));
            bool recordingMissing = !string.IsNullOrEmpty(profile.RecordingDeviceId)
                && !allRecording.Any(d => string.Equals(d.Id, profile.RecordingDeviceId, StringComparison.Ordinal));
            var missingNotes = new List<string>();
            if (playbackMissing) missingNotes.Add($"播放: {profile.PlaybackDeviceName ?? "未知"} (未连接)");
            if (recordingMissing) missingNotes.Add($"录音: {profile.RecordingDeviceName ?? "未知"} (未连接)");

            bool isLocked = SettingsService.Load().LockedProfileId == profile.Id;

            var panel = new StackPanel { Orientation = Orientation.Horizontal };
            if (!string.IsNullOrEmpty(profile.Color))
            {
                try
                {
                    var c = (Color)ColorConverter.ConvertFromString(profile.Color);
                    panel.Children.Add(new System.Windows.Shapes.Ellipse
                    {
                        Width = 8,
                        Height = 8,
                        Fill = new SolidColorBrush(c),
                        Margin = new Thickness(0, 0, 6, 0),
                        VerticalAlignment = VerticalAlignment.Center,
                    });
                }
                catch { }
            }
            if (isLocked)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = "\U0001F512 ",
                    Foreground = new SolidColorBrush(Color.FromRgb(0xD9, 0x77, 0x06)),
                    VerticalAlignment = VerticalAlignment.Center,
                });
            }
            if (isActive)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = "\u2714 ",
                    Foreground = new SolidColorBrush(Color.FromRgb(0x10, 0x9E, 0x7A)),
                    FontWeight = FontWeights.SemiBold,
                    VerticalAlignment = VerticalAlignment.Center,
                });
            }
            if (missingNotes.Count > 0)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = "⚠ ",
                    Foreground = new SolidColorBrush(Color.FromRgb(0xD3, 0x2F, 0x2F)),
                    FontWeight = FontWeights.SemiBold,
                    VerticalAlignment = VerticalAlignment.Center,
                });
            }
            panel.Children.Add(new TextBlock
            {
                Text = profile.Name,
                VerticalAlignment = VerticalAlignment.Center,
            });

            var btn = new Button
            {
                Content = panel,
                Tag = profile.Id,
                Style = (Style)FindResource(isActive ? "MiniActiveProfileButton" : "MiniProfileButton"),
                HorizontalContentAlignment = HorizontalAlignment.Left,
            };
            var tipParts = new List<string>();
            if (nicknameLine != null) tipParts.Add(nicknameLine);
            if (missingNotes.Count > 0) tipParts.Add(string.Join("\n", missingNotes));
            if (tipParts.Count > 0) btn.ToolTip = string.Join("\n\n", tipParts);
            btn.Click += ProfileButton_Click;
            ProfileList.Items.Add(btn);
        }

        // Bluetooth warning
        bool hasBluetooth = AudioDeviceService.HasBluetoothDevice();
        BluetoothWarning.Visibility = hasBluetooth ? Visibility.Visible : Visibility.Collapsed;

        // Profile mismatch warning — consider hidden profiles too so hiding the
        // currently-active one doesn't make the warning blink.
        bool anyActiveAcrossAll = allProfiles.Any(p =>
            p.PlaybackDeviceId == currentPlayback?.Id
            && p.RecordingDeviceId == currentRecording?.Id
            && p.PlaybackDeviceId == currentPlaybackComm.Id
            && p.RecordingDeviceId == currentRecordingComm.Id);
        bool showWarning = allProfiles.Count > 0 && !anyActiveAcrossAll;
        WarningText.Visibility = showWarning ? Visibility.Visible : Visibility.Collapsed;
        WarningText.Tag = (showWarning && SettingsService.Load().EnableBlinkAnimation) ? "Blink" : null;

        var drifted = ((App)Application.Current).DriftedApps;
        if (drifted.Count > 0)
        {
            AppDriftWarning.Text = $"{drifted.Count} \u4E2A\u5E94\u7528\u504F\u79BB";
            AppDriftWarning.ToolTip = string.Join(", ", drifted);
            AppDriftWarning.Visibility = Visibility.Visible;
        }
        else
        {
            AppDriftWarning.Visibility = Visibility.Collapsed;
        }

        // Border color tint: red if profile mismatch, amber if bluetooth only, neutral otherwise
        _border.BorderBrush = showWarning
            ? new SolidColorBrush(Color.FromRgb(0xFC, 0xA5, 0xA5))
            : hasBluetooth
                ? new SolidColorBrush(Color.FromRgb(0xFC, 0xD3, 0x4D))
                : new SolidColorBrush(Color.FromRgb(0xE5, 0xE7, 0xEB));
        _border.BorderThickness = new Thickness(1);
        _border.Background = new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF));
    }

    private void ProfileButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isLocked) return;
        if (sender is not Button btn || btn.Tag is not Guid id) return;

        var profile = ProfileService.GetAll().Find(p => p.Id == id);
        if (profile == null) return;

        if (!((App)Application.Current).TryUserApplyProfile(profile)) return;

        try
        {
            var result = ProfileApplyService.Apply(profile);
            ((App)Application.Current).MarkOwnChange();

            LoadProfiles();
            _onProfileApplied?.Invoke();

            ((App)Application.Current).NotifyProfileApplied(profile, result);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"\u5E94\u7528\u914D\u7F6E\u5931\u8D25\uFF1A\n{ex.Message}", "\u9519\u8BEF",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Lock_Changed(object sender, RoutedEventArgs e)
    {
        _isLocked = LockCheck.IsChecked == true;
        if (ProfileList != null)
            ProfileList.Opacity = _isLocked ? 0.5 : 1.0;
    }

    private void Opacity_Changed(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        Opacity = OpacitySlider.Value;
        SaveSettings();
    }

    private void Window_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.OriginalSource is DependencyObject source
            && (FindParent<Button>(source) != null
                || FindParent<CheckBox>(source) != null
                || FindParent<Slider>(source) != null))
            return;

        DragMove();
    }

    private static T? FindParent<T>(DependencyObject child) where T : DependencyObject
    {
        var parent = VisualTreeHelper.GetParent(child);
        while (parent != null)
        {
            if (parent is T found) return found;
            parent = VisualTreeHelper.GetParent(parent);
        }
        return null;
    }
}
