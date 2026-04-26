using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using NAudio.CoreAudioApi;

namespace AudioDeviceSwitcher;

public enum VoicemeeterRestartStatus
{
    NotRequested,
    Restarted,
    NotInstalled,
    NotRunning,
    Failed,
}

public record VoicemeeterIoSlot(string Label, string? CustomLabel, string? DeviceName, bool Muted = false, int Index = 0, bool DeviceMissing = false);

public record VoicemeeterIoState(string TypeName, List<VoicemeeterIoSlot> Outputs, List<VoicemeeterIoSlot> Inputs);

// Voicemeeter Remote API wrapper. Loads VoicemeeterRemote64.dll from the install
// directory found via the uninstall registry key, then keeps a long-lived Login
// session so VBVMR_IsParametersDirty can be polled to detect changes.
public static partial class VoicemeeterService
{
    private const string DllName = "VoicemeeterRemote64.dll";
    private const string DllName32 = "VoicemeeterRemote.dll";

    private static bool _loaded;
    private static bool _loadAttempted;
    private static bool _loggedIn;
    private static readonly object _gate = new();

    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "AudioDeviceSwitcher", "voicemeeter.log");

    private static void Log(string msg)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
            File.AppendAllText(LogPath, $"[{DateTime.Now:HH:mm:ss.fff}] {msg}\n");
        }
        catch { }
    }

    public static bool EnsureLoaded()
    {
        // User opted out of Voicemeeter integration — never attempt to load the DLL.
        if (!SettingsService.Load().VoicemeeterIntegrationEnabled) return false;
        if (_loaded) return true;
        if (_loadAttempted) return false;
        _loadAttempted = true;

        var dir = FindInstallDir();
        if (dir == null) { Log("install dir not found in registry"); return false; }
        Log($"install dir: {dir}");

        var preferred = Environment.Is64BitProcess ? DllName : DllName32;
        var fallback = Environment.Is64BitProcess ? DllName32 : DllName;
        var fullPath = Path.Combine(dir, preferred);
        if (!File.Exists(fullPath)) fullPath = Path.Combine(dir, fallback);
        if (!File.Exists(fullPath)) { Log($"dll not found in {dir}"); return false; }
        Log($"loading dll: {fullPath}");

        var h = LoadLibrary(fullPath);
        if (h == IntPtr.Zero) { Log($"LoadLibrary failed, GetLastError={Marshal.GetLastWin32Error()}"); return false; }

        _loaded = true;
        return true;
    }

    private static bool EnsureLoggedIn()
    {
        if (_loggedIn) return true;
        if (!EnsureLoaded()) return false;

        int rc = VBVMR_Login();
        Log($"Login rc={rc}");
        // 0 = ok, 1 = already logged in (still fine); negatives are errors.
        if (rc < 0) return false;

        // SDK requires waiting after login before issuing any other calls.
        Thread.Sleep(100);
        _loggedIn = true;
        return true;
    }

    private static void DropSession()
    {
        if (!_loggedIn) return;
        try { VBVMR_Logout(); } catch { }
        _loggedIn = false;
    }

    public static void Shutdown()
    {
        lock (_gate) { DropSession(); }
    }

    // Called when the user toggles VoicemeeterIntegrationEnabled in settings.
    // Drops any active session so re-enabling later starts clean.
    public static void ResetLoadState()
    {
        lock (_gate)
        {
            DropSession();
            _loadAttempted = false;
        }
    }

    public static VoicemeeterRestartStatus RestartAudioEngine()
    {
        lock (_gate)
        {
            if (!EnsureLoaded()) return VoicemeeterRestartStatus.NotInstalled;
            if (!EnsureLoggedIn()) return VoicemeeterRestartStatus.Failed;
            try
            {
                if (VBVMR_GetVoicemeeterVersion(out _) < 0)
                {
                    DropSession();
                    return VoicemeeterRestartStatus.NotRunning;
                }
                int setRc = VBVMR_SetParameterFloat("Command.Restart", 1.0f);
                Log($"Command.Restart rc={setRc}");
                if (setRc < 0) return VoicemeeterRestartStatus.Failed;
                Thread.Sleep(200);
                return VoicemeeterRestartStatus.Restarted;
            }
            catch (Exception ex) { Log($"RestartAudioEngine exception: {ex.Message}"); DropSession(); return VoicemeeterRestartStatus.Failed; }
        }
    }

    public static bool IsParametersDirty()
    {
        lock (_gate)
        {
            if (!EnsureLoaded()) return false;
            if (!EnsureLoggedIn()) return false;
            try
            {
                int rc = VBVMR_IsParametersDirty();
                if (rc < 0) { DropSession(); return false; }
                return rc == 1;
            }
            catch { DropSession(); return false; }
        }
    }

    public static VoicemeeterIoState? GetIoState()
    {
        lock (_gate)
        {
            if (!EnsureLoaded()) return null;
            if (!EnsureLoggedIn()) return null;
            try
            {
                if (VBVMR_GetVoicemeeterType(out int type) < 0 || type <= 0)
                {
                    DropSession();
                    return null;
                }

                // Voicemeeter editions:
                //   1 = Standard (1 hardware out, 2 hardware in)
                //   2 = Banana   (3 hardware out, 3 hardware in)
                //   3 = Potato   (5 hardware out, 5 hardware in)
                (string name, int outCount, int inCount) = type switch
                {
                    1 => ("Voicemeeter", 1, 2),
                    2 => ("Voicemeeter Banana", 3, 3),
                    3 => ("Voicemeeter Potato", 5, 5),
                    _ => ($"Voicemeeter (type {type})", 3, 3),
                };

                // Enumerate system audio endpoints once so we can flag any Strip/Bus
                // pointing at a device that's no longer plugged in (Voicemeeter shows red).
                var systemNames = EnumerateSystemDeviceNames();

                var outputs = new List<VoicemeeterIoSlot>();
                for (int i = 0; i < outCount; i++)
                {
                    var dn = ReadStringParam($"Bus[{i}].device.name");
                    outputs.Add(new VoicemeeterIoSlot(
                        $"A{i + 1}",
                        ReadStringParam($"Bus[{i}].label"),
                        dn,
                        Index: i,
                        DeviceMissing: !string.IsNullOrEmpty(dn) && !DeviceMatches(dn, systemNames)));
                }

                var inputs = new List<VoicemeeterIoSlot>();
                for (int i = 0; i < inCount; i++)
                {
                    var dn = ReadStringParam($"Strip[{i}].device.name");
                    inputs.Add(new VoicemeeterIoSlot(
                        $"Strip {i + 1}",
                        ReadStringParam($"Strip[{i}].label"),
                        dn,
                        ReadFloatParam($"Strip[{i}].Mute") >= 0.5f,
                        Index: i,
                        DeviceMissing: !string.IsNullOrEmpty(dn) && !DeviceMatches(dn, systemNames)));
                }

                return new VoicemeeterIoState(name, outputs, inputs);
            }
            catch (Exception ex) { Log($"GetIoState exception: {ex.Message}"); DropSession(); return null; }
        }
    }

    private static float ReadFloatParam(string param)
    {
        int rc = VBVMR_GetParameterFloat(param, out float v);
        return rc < 0 ? 0f : v;
    }

    // Peak levels for displayed strips (post-mute, 2 ch each) and buses (8 ch each, take max).
    public static (float[] Outputs, float[] Inputs) GetPeakLevels(int outCount, int inCount)
    {
        lock (_gate)
        {
            if (!_loggedIn || !_loaded) return ([], []);
            var outputs = new float[outCount];
            for (int i = 0; i < outCount; i++)
            {
                float max = 0f;
                for (int ch = 0; ch < 8; ch++)
                {
                    if (VBVMR_GetLevel(3, i * 8 + ch, out float v) >= 0 && v > max) max = v;
                }
                outputs[i] = max;
            }
            var inputs = new float[inCount];
            for (int i = 0; i < inCount; i++)
            {
                float max = 0f;
                for (int ch = 0; ch < 2; ch++)
                {
                    if (VBVMR_GetLevel(2, i * 2 + ch, out float v) >= 0 && v > max) max = v;
                }
                inputs[i] = max;
            }
            return (outputs, inputs);
        }
    }

    public static bool SetStripMute(int stripIndex, bool muted)
    {
        lock (_gate)
        {
            if (!EnsureLoggedIn()) return false;
            try
            {
                int rc = VBVMR_SetParameterFloat($"Strip[{stripIndex}].Mute", muted ? 1f : 0f);
                return rc >= 0;
            }
            catch { return false; }
        }
    }

    private static List<string> EnumerateSystemDeviceNames()
    {
        var names = new List<string>();
        try
        {
            using var enumerator = new MMDeviceEnumerator();
            // Active only — disabled / unplugged devices won't satisfy a Voicemeeter binding either.
            var devs = enumerator.EnumerateAudioEndPoints(DataFlow.All, DeviceState.Active);
            foreach (var d in devs)
            {
                try { names.Add(d.FriendlyName); }
                catch { }
                finally { d.Dispose(); }
            }
        }
        catch { }
        return names;
    }

    private static bool DeviceMatches(string vmName, List<string> systemNames)
    {
        // Voicemeeter MME names get truncated (~31 chars) so do prefix matching both ways.
        foreach (var s in systemNames)
        {
            if (s.Equals(vmName, StringComparison.OrdinalIgnoreCase)) return true;
            if (s.StartsWith(vmName, StringComparison.OrdinalIgnoreCase)) return true;
            if (vmName.StartsWith(s, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private static string? ReadStringParam(string param)
    {
        // Voicemeeter param NAMES are always ASCII ("Strip[0].label" etc.) but the
        // returned VALUE can contain Unicode (custom labels, device names with CJK
        // chars). The W variant returns UTF-16 directly — the A variant returns
        // ANSI-encoded text that mojibakes for non-ASCII.
        var sb = new StringBuilder(512);
        int rc = VBVMR_GetParameterStringW(param, sb);
        if (rc < 0) return null;
        var s = sb.ToString();
        return string.IsNullOrWhiteSpace(s) ? null : s;
    }

    private static string? FindInstallDir()
    {
        // Voicemeeter Standard / Banana / Potato each register under a different GUID
        // but the display name always starts with "VB:Voicemeeter".
        string[] roots =
        [
            @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        ];
        foreach (var root in roots)
        {
            try
            {
                using var rk = Registry.LocalMachine.OpenSubKey(root);
                if (rk == null) continue;
                foreach (var name in rk.GetSubKeyNames())
                {
                    if (!name.StartsWith("VB:Voicemeeter", StringComparison.OrdinalIgnoreCase)) continue;
                    using var sub = rk.OpenSubKey(name);
                    if (sub == null) continue;
                    var dir = ExtractDir(sub.GetValue("InstallLocation") as string)
                           ?? ExtractDir(sub.GetValue("UninstallString") as string);
                    if (dir != null) return dir;
                }
            }
            catch { }
        }
        return null;
    }

    private static string? ExtractDir(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        var s = raw.Trim().Trim('"');
        try
        {
            if (Directory.Exists(s)) return s;
            var dir = Path.GetDirectoryName(s);
            if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir)) return dir;
        }
        catch { }
        return null;
    }

    [LibraryImport("kernel32.dll", EntryPoint = "LoadLibraryW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    private static partial IntPtr LoadLibrary(string path);

    [DllImport(DllName, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_Login();

    [DllImport(DllName, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_Logout();

    [DllImport(DllName, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_GetVoicemeeterVersion(out int version);

    [DllImport(DllName, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_GetVoicemeeterType(out int type);

    [DllImport(DllName, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_IsParametersDirty();

    [DllImport(DllName, CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_SetParameterFloat(string param, float value);

    [DllImport(DllName, CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_GetParameterStringA(string param, StringBuilder buf);

    [DllImport(DllName, CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_GetParameterFloat(string param, out float value);

    [DllImport(DllName, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_GetLevel(int nType, int nuChannel, out float pValue);

    [DllImport(DllName, CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_GetParameterStringW(
        [MarshalAs(UnmanagedType.LPStr)] string param,
        StringBuilder buf);
}
