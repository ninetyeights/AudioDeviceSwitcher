using System.IO;
using System.Text.Json;

namespace AudioDeviceSwitcher;

public class AppSettings
{
    public double? MainWindowLeft { get; set; }
    public double? MainWindowTop { get; set; }
    public double? MainWindowWidth { get; set; }
    public double? MainWindowHeight { get; set; }
    public double? MiniWindowLeft { get; set; }
    public double? MiniWindowTop { get; set; }
    public double MiniWindowOpacity { get; set; } = 1.0;
    public bool MiniWindowVisible { get; set; }
    public List<string> HiddenDeviceIds { get; set; } = [];
    public bool ShowHiddenDevices { get; set; }
    public bool ShowDisabledDevices { get; set; }
    public Dictionary<string, string> DeviceNicknames { get; set; } = [];

    public bool NotifyProfileApplied { get; set; } = true;
    public bool NotifyDeviceChanged { get; set; } = true;
    public bool NotifyBluetooth { get; set; } = true;
    public bool NotifyAppDrift { get; set; } = true;
    public bool EnableBlinkAnimation { get; set; } = true;
    public bool StartMinimized { get; set; } = false;
    public Guid? LockedProfileId { get; set; }
    public bool VoicemeeterMuteLocked { get; set; }
    public List<bool> VoicemeeterStripMuteSnapshot { get; set; } = [];
    public bool VoicemeeterIntegrationEnabled { get; set; } = false;
}

public static class SettingsService
{
    private static readonly string SettingsDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AudioDeviceSwitcher");

    private static readonly string SettingsPath = Path.Combine(SettingsDir, "settings.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private static AppSettings? _cached;

    public static AppSettings Load()
    {
        if (_cached != null) return _cached;

        if (!File.Exists(SettingsPath))
        {
            _cached = new AppSettings();
            return _cached;
        }

        try
        {
            var json = File.ReadAllText(SettingsPath);
            _cached = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
        }
        catch
        {
            _cached = new AppSettings();
        }
        return _cached;
    }

    public static void Save()
    {
        if (_cached == null) return;
        Directory.CreateDirectory(SettingsDir);
        var json = JsonSerializer.Serialize(_cached, JsonOptions);
        File.WriteAllText(SettingsPath, json);
    }
}
