using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

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
            }
        };
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
        // Skip if hidden — nothing to show anyway
        if (!IsVisible) return;

        // Skip if state hasn't changed (avoid needless UI rebuild)
        var signature = ComputeStateSignature();
        if (signature == _lastStateSignature) return;
        _lastStateSignature = signature;

        LoadDevices();
        LoadProfiles();
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
        LoadPlaybackDevices();
        LoadRecordingDevices();
        BluetoothWarning.Visibility = AudioDeviceService.HasBluetoothDevice()
            ? Visibility.Visible : Visibility.Collapsed;
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
            var btn = CreateDeviceButton(device, isHidden, isSelected);
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
            var btn = CreateDeviceButton(device, isHidden, isSelected);
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

    private Button CreateDeviceButton(AudioDeviceInfo device, bool isHidden = false, bool isSelected = false)
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

        var btn = new Button
        {
            Content = root,
            Tag = device.Id,
            Style = style,
            Opacity = opacity,
            Cursor = device.IsDisabled ? System.Windows.Input.Cursors.Arrow : System.Windows.Input.Cursors.Hand,
        };
        btn.Click += DeviceButton_Click;

        var menu = new ContextMenu();
        var setDefaultItem = new MenuItem { Header = "设为默认设备", Tag = device.Id };
        setDefaultItem.Click += SetDefaultDevice_Click;
        setDefaultItem.IsEnabled = !device.IsDefault && !device.IsDisabled;
        menu.Items.Add(setDefaultItem);
        var toggleEnableItem = new MenuItem
        {
            Header = device.IsDisabled ? "启用此设备" : "禁用此设备",
            Tag = device.Id,
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
        var currentPlayback = AudioDeviceService.GetPlaybackDevices().Find(d => d.IsDefault);
        var currentRecording = AudioDeviceService.GetRecordingDevices().Find(d => d.IsDefault);
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

            // Left side: name + device info
            var nameText = new TextBlock
            {
                Text = (isActive ? "\u2714 " : "") + profile.Name,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = isActive
                    ? new SolidColorBrush(Color.FromRgb(0x1E, 0x40, 0xAF))
                    : new SolidColorBrush(Color.FromRgb(0x11, 0x18, 0x27)),
            };

            var detailParts = $"\u25B6 {ResolveDeviceName(profile.PlaybackDeviceId, profile.PlaybackDeviceName)}  \u00B7  \u25CF {ResolveDeviceName(profile.RecordingDeviceId, profile.RecordingDeviceName)}";
            if (profile.HotkeyKey != 0)
                detailParts += $"  \u00B7  \u2328 {FormatHotkey(profile.HotkeyModifiers, profile.HotkeyKey)}";

            var detailText = new TextBlock
            {
                Text = detailParts,
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80)),
                Margin = new Thickness(0, 3, 0, 0),
                TextTrimming = TextTrimming.CharacterEllipsis,
            };

            var infoPanel = new StackPanel();
            infoPanel.Children.Add(nameText);
            infoPanel.Children.Add(detailText);

            var applyBtn = new Button
            {
                Content = infoPanel,
                Style = (Style)FindResource("ProfileApplyButton"),
                Tag = profile.Id,
            };
            applyBtn.Click += ApplyProfile_Click;

            // Right side: edit + delete links
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
            actionPanel.Children.Add(editBtn);
            actionPanel.Children.Add(separator);
            actionPanel.Children.Add(deleteBtn);

            // Card layout
            var dock = new DockPanel();
            DockPanel.SetDock(actionPanel, Dock.Right);
            dock.Children.Add(actionPanel);
            dock.Children.Add(applyBtn);

            var card = new Border
            {
                Style = (Style)FindResource(isActive ? "ActiveProfileCard" : "ProfileCard"),
                Child = dock,
                Cursor = System.Windows.Input.Cursors.Hand,
            };
            card.MouseLeftButtonUp += (_, e) =>
            {
                if (e.ChangedButton != System.Windows.Input.MouseButton.Left) return;
                ApplyProfile_Click(applyBtn, e);
            };

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

    private void ApplyProfile_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not Guid id) return;

        var profile = ProfileService.GetAll().Find(p => p.Id == id);
        if (profile == null) return;

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

        var dialog = new ProfileEditDialog(profile.Name, profile.HotkeyModifiers, profile.HotkeyKey, profile.AppOverrides) { Owner = this };
        if (dialog.ShowDialog() != true) return;

        profile.Name = dialog.ProfileName;
        profile.HotkeyModifiers = dialog.HotkeyModifiers;
        profile.HotkeyKey = dialog.HotkeyKey;
        profile.AppOverrides = dialog.AppOverrides;
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
            RegisterProfileHotkeys();
            LoadProfiles();
            _miniWindow?.LoadProfiles();
        }
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
