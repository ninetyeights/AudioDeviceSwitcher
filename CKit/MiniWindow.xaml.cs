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

        var profiles = ProfileService.GetAll();
        EmptyHint.Visibility = profiles.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

        var currentPlayback = AudioDeviceService.GetPlaybackDevices().Find(d => d.IsDefault);
        var currentRecording = AudioDeviceService.GetRecordingDevices().Find(d => d.IsDefault);
        var currentPlaybackComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Render);
        var currentRecordingComm = AudioDeviceService.GetCommunicationsDefault(NAudio.CoreAudioApi.DataFlow.Capture);

        var nicknames = SettingsService.Load().DeviceNicknames;
        string? NicknameOf(string? id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            return nicknames.TryGetValue(id, out var nn) && !string.IsNullOrEmpty(nn) ? nn : null;
        }

        bool anyActive = false;

        foreach (var profile in profiles)
        {
            bool isActive = profile.PlaybackDeviceId == currentPlayback?.Id
                         && profile.RecordingDeviceId == currentRecording?.Id
                         && profile.PlaybackDeviceId == currentPlaybackComm.Id
                         && profile.RecordingDeviceId == currentRecordingComm.Id;
            if (isActive) anyActive = true;

            var playNn = NicknameOf(profile.PlaybackDeviceId);
            var recNn = NicknameOf(profile.RecordingDeviceId);
            var nicknameParts = new List<string>();
            if (playNn != null) nicknameParts.Add($"\u25B6 {playNn}");
            if (recNn != null) nicknameParts.Add($"\u25CF {recNn}");
            string? nicknameLine = nicknameParts.Count > 0 ? string.Join("  ", nicknameParts) : null;

            var panel = new StackPanel { Orientation = Orientation.Horizontal };
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
            if (nicknameLine != null)
                btn.ToolTip = nicknameLine;
            btn.Click += ProfileButton_Click;
            ProfileList.Items.Add(btn);
        }

        // Bluetooth warning
        bool hasBluetooth = AudioDeviceService.HasBluetoothDevice();
        BluetoothWarning.Visibility = hasBluetooth ? Visibility.Visible : Visibility.Collapsed;

        // Profile mismatch warning
        bool showWarning = profiles.Count > 0 && !anyActive;
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
