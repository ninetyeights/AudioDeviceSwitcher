using System.Diagnostics;
using System.IO;
using System.Windows;

namespace AudioDeviceSwitcher;

public partial class SettingsWindow : Window
{
    public SettingsWindow()
    {
        InitializeComponent();
        var s = SettingsService.Load();
        ChkNotifyProfile.IsChecked = s.NotifyProfileApplied;
        ChkNotifyDevice.IsChecked = s.NotifyDeviceChanged;
        ChkNotifyBluetooth.IsChecked = s.NotifyBluetooth;
        ChkNotifyDrift.IsChecked = s.NotifyAppDrift;
        ChkBlink.IsChecked = s.EnableBlinkAnimation;
        ChkStartMinimized.IsChecked = s.StartMinimized;
        ChkVoicemeeterEnabled.IsChecked = s.VoicemeeterIntegrationEnabled;
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        var s = SettingsService.Load();
        bool prevVm = s.VoicemeeterIntegrationEnabled;
        s.NotifyProfileApplied = ChkNotifyProfile.IsChecked == true;
        s.NotifyDeviceChanged = ChkNotifyDevice.IsChecked == true;
        s.NotifyBluetooth = ChkNotifyBluetooth.IsChecked == true;
        s.NotifyAppDrift = ChkNotifyDrift.IsChecked == true;
        s.EnableBlinkAnimation = ChkBlink.IsChecked == true;
        s.StartMinimized = ChkStartMinimized.IsChecked == true;
        s.VoicemeeterIntegrationEnabled = ChkVoicemeeterEnabled.IsChecked == true;
        SettingsService.Save();
        if (prevVm != s.VoicemeeterIntegrationEnabled)
        {
            // Drop any cached login/load state so the new setting takes effect immediately.
            VoicemeeterService.ResetLoadState();
        }
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void OpenDataDir_Click(object sender, RoutedEventArgs e)
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AudioDeviceSwitcher");
        Directory.CreateDirectory(dir);
        Process.Start(new ProcessStartInfo("explorer.exe", dir) { UseShellExecute = false });
    }
}
