using System.IO;
using System.Text.Json;

namespace AudioDeviceSwitcher;

public record DeviceProfile
{
    public Guid Id { get; init; } = Guid.NewGuid();
    public string Name { get; set; } = "";
    public string? PlaybackDeviceId { get; set; }
    public string? PlaybackDeviceName { get; set; }
    public string? RecordingDeviceId { get; set; }
    public string? RecordingDeviceName { get; set; }
    public int HotkeyModifiers { get; set; }
    public int HotkeyKey { get; set; }
    public List<AppOverride> AppOverrides { get; set; } = [];
}

public record AppOverride
{
    public string ExePath { get; set; } = "";
    public Guid AppProfileId { get; set; }
}

public static class ProfileService
{
    private static readonly string ProfileDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AudioDeviceSwitcher");

    private static readonly string ProfilePath = Path.Combine(ProfileDir, "profiles.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public static List<DeviceProfile> GetAll()
    {
        if (!File.Exists(ProfilePath))
            return [];

        var json = File.ReadAllText(ProfilePath);
        return JsonSerializer.Deserialize<List<DeviceProfile>>(json, JsonOptions) ?? [];
    }

    public static void Save(DeviceProfile profile)
    {
        var profiles = GetAll();
        var index = profiles.FindIndex(p => p.Id == profile.Id);
        if (index >= 0)
            profiles[index] = profile;
        else
            profiles.Add(profile);

        WriteAll(profiles);
    }

    public static void Delete(Guid id)
    {
        var profiles = GetAll();
        profiles.RemoveAll(p => p.Id == id);
        WriteAll(profiles);
    }

    private static void WriteAll(List<DeviceProfile> profiles)
    {
        Directory.CreateDirectory(ProfileDir);
        var json = JsonSerializer.Serialize(profiles, JsonOptions);
        File.WriteAllText(ProfilePath, json);
    }
}
