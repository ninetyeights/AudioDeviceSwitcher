using Microsoft.Win32;

namespace AudioDeviceSwitcher;

public static class AutoStartService
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "AudioDeviceSwitcher";

    public static bool IsEnabled()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
            if (key == null) return false;
            var value = key.GetValue(ValueName) as string;
            if (string.IsNullOrEmpty(value)) return false;

            var exe = Environment.ProcessPath;
            if (string.IsNullOrEmpty(exe)) return true;
            // Value stored quoted; normalize for compare
            var stored = value.Trim().Trim('"');
            return string.Equals(stored, exe, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    public static void SetEnabled(bool enabled)
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, true)
                        ?? Registry.CurrentUser.CreateSubKey(RunKeyPath);
        if (key == null) return;

        if (enabled)
        {
            var exe = Environment.ProcessPath;
            if (string.IsNullOrEmpty(exe)) return;
            key.SetValue(ValueName, $"\"{exe}\"", RegistryValueKind.String);
        }
        else
        {
            if (key.GetValue(ValueName) != null)
                key.DeleteValue(ValueName, false);
        }
    }
}
