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

public record VoicemeeterIoSlot(string Label, string? CustomLabel, string? DeviceName, bool Muted = false, int Index = 0, bool DeviceMissing = false, bool IsVirtual = false);

public record VoicemeeterIoState(string TypeName, List<VoicemeeterIoSlot> Outputs, List<VoicemeeterIoSlot> Inputs);

// A device the user can pick for a Strip/Bus, as Voicemeeter itself enumerates it.
// Driver is the param suffix ("wdm"/"mme"/"ks"/"asio"); DriverLabel is its display form.
// Name is the exact string Voicemeeter expects when selecting it (NEVER substitute the
// Windows name here — the API matches on it). DisplayName is the menu label, enriched
// with the Windows friendly name when Voicemeeter only exposes a generic KS name
// ("Headphones") or a truncated MME name; falls back to Name when no match is found.
public record VoicemeeterDevice(int Type, string Driver, string DriverLabel, string Name, string DisplayName);

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

    private static readonly object _logGate = new();
    private const long LogMaxBytes = 512 * 1024; // rotate past 512 KB; keep one .old

    private static void Log(string msg)
    {
        try
        {
            // Logging isn't on a hot path (event-driven, not per-poll); serialize so
            // concurrent writers don't interleave or fight over the file handle.
            lock (_logGate)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
                var fi = new FileInfo(LogPath);
                if (fi.Exists && fi.Length > LogMaxBytes)
                {
                    var old = LogPath + ".old";
                    File.Delete(old);          // no-op if absent
                    File.Move(LogPath, old);   // current -> .old, fresh file starts below
                }
                File.AppendAllText(LogPath, $"[{DateTime.Now:HH:mm:ss.fff}] {msg}\n");
            }
        }
        catch { }
    }

    // Lets the lock-enforcement code (in MainWindow) write to the same diagnostic log.
    public static void LogExternal(string msg) => Log(msg);

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
            InvalidateSysSnapshot();
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

                // Voicemeeter editions — hardware in/out plus the trailing VIRTUAL input
                // strips (Voicemeeter VAIO / AUX / VAIO3), which have no selectable device:
                //   1 = Standard (1 hw out, 2 hw in, 1 virtual in)
                //   2 = Banana   (3 hw out, 3 hw in, 2 virtual in)
                //   3 = Potato   (5 hw out, 5 hw in, 3 virtual in)
                (string name, int outCount, int inCount, int virtualInCount) = type switch
                {
                    1 => ("Voicemeeter", 1, 2, 1),
                    2 => ("Voicemeeter Banana", 3, 3, 2),
                    3 => ("Voicemeeter Potato", 5, 5, 3),
                    _ => ($"Voicemeeter (type {type})", 3, 3, 2),
                };

                // System endpoints + the name-enrichment map (cached, rebuilt every few
                // seconds): used to (a) show the current device's model name and (b) decide
                // "missing". Match Voicemeeter's leniency — a configured device is "missing"
                // (red) ONLY when gone from the system entirely, not when a paired device is
                // merely disconnected (a connected-but-idle BT headphone reports Unplugged,
                // yet Voicemeeter still shows it normally), so any enumerated state counts as
                // present. The per-poll device.name reads below stay fresh for drift detection.
                var snap = GetSysSnapshot();
                var systemNames = snap.Names;
                var displayMap = snap.DisplayMap;

                // Check the ENRICHED name: "Headphones [A]" never matches the system's
                // "Headphones (HD 450BT)", so the raw name would falsely flag it as missing.
                bool Missing(string? dn, bool isInput)
                {
                    if (string.IsNullOrEmpty(dn)) return false;
                    var resolved = displayMap.TryGetValue($"{(isInput ? "I" : "O")}|{dn}", out var d) ? d : dn;
                    return !DeviceMatches(resolved, systemNames);
                }

                var outputs = new List<VoicemeeterIoSlot>();
                for (int i = 0; i < outCount; i++)
                {
                    var dn = ReadStringParam($"Bus[{i}].device.name");
                    outputs.Add(new VoicemeeterIoSlot(
                        $"A{i + 1}",
                        ReadStringParam($"Bus[{i}].label"),
                        dn,
                        Index: i,
                        DeviceMissing: Missing(dn, false)));
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
                        DeviceMissing: Missing(dn, true)));
                }

                // Virtual input strips follow the hardware ones. They have no selectable
                // hardware device — show the virtual input's name (static) + mute only.
                string[] virtualNames = ["Voicemeeter VAIO", "Voicemeeter AUX", "Voicemeeter VAIO3"];
                for (int v = 0; v < virtualInCount; v++)
                {
                    int idx = inCount + v;
                    inputs.Add(new VoicemeeterIoSlot(
                        $"Strip {idx + 1}",
                        ReadStringParam($"Strip[{idx}].label"),
                        v < virtualNames.Length ? virtualNames[v] : $"Voicemeeter Virtual {v + 1}",
                        ReadFloatParam($"Strip[{idx}].Mute") >= 0.5f,
                        Index: idx,
                        IsVirtual: true));
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
    // Called every 50ms on the UI thread, so it must never block: if another VM call
    // (device set/enumerate, GetIoState) is holding the gate, skip this frame instead
    // of freezing the UI. The bars just stall for a moment — cosmetic only.
    public static (float[] Outputs, float[] Inputs) GetPeakLevels(int outCount, int inCount)
    {
        if (!Monitor.TryEnter(_gate)) return ([], []);
        try
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
        finally { Monitor.Exit(_gate); }
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

    // Voicemeeter device-type ids (from VoicemeeterRemote.h) and the param suffix used
    // to select a device of that type: Bus[i].device.wdm = "<name>" etc.
    private static string? DriverSuffix(int type) => type switch
    {
        1 => "mme",
        3 => "wdm",
        4 => "ks",
        5 => "asio",
        _ => null,
    };

    // Prefer WDM (Voicemeeter's modern default), then KS, MME, ASIO when the same
    // device name is enumerated under multiple drivers.
    private static int DriverPriority(int type) => type switch
    {
        3 => 0, // wdm
        4 => 1, // ks
        1 => 2, // mme
        5 => 3, // asio
        _ => 9,
    };

    private static List<(int Type, string Name, string Hwid)> EnumVmDevices(bool isInput)
    {
        var list = new List<(int, string, string)>();
        int n = isInput ? VBVMR_Input_GetDeviceNumber() : VBVMR_Output_GetDeviceNumber();
        for (int i = 0; i < n; i++)
        {
            var name = new StringBuilder(512);
            var hwid = new StringBuilder(512);
            int rc = isInput
                ? VBVMR_Input_GetDeviceDescW(i, out int t, name, hwid)
                : VBVMR_Output_GetDeviceDescW(i, out t, name, hwid);
            if (rc < 0) continue;
            var nm = name.ToString();
            if (!string.IsNullOrWhiteSpace(nm)) list.Add((t, nm, hwid.ToString()));
        }
        return list;
    }

    // Restore a Strip/Bus to its locked device. Voicemeeter selects a device by the
    // name from ITS OWN enumeration (VBVMR_*_GetDeviceDescW) paired with that device's
    // driver type. Returns false ONLY when the device genuinely isn't enumerated
    // (truly gone) or Voicemeeter rejects the set — those are the real "give up"
    // cases. rc>=0 with an enumerated match is trusted as applied: the device.name
    // readback lags the swap by >1s, so verifying it here produced false failures
    // that tripped the give-up and disabled the lock after one restore.
    public static bool RestoreIoDevice(bool isInput, int index, string targetName)
    {
        if (string.IsNullOrEmpty(targetName)) return false;
        lock (_gate)
        {
            if (!EnsureLoggedIn()) return false;
            string slot = isInput ? $"Strip[{index}]" : $"Bus[{index}]";
            try
            {
                var current = ReadStringParam($"{slot}.device.name");
                if (NamesMatch(current, targetName)) return true;

                var devs = EnumVmDevices(isInput);
                (int Type, string Name)? best = null;
                foreach (var d in devs)
                {
                    if (DriverSuffix(d.Type) == null) continue;
                    if (!NamesMatch(d.Name, targetName)) continue;
                    if (best == null || DriverPriority(d.Type) < DriverPriority(best.Value.Type))
                        best = (d.Type, d.Name);
                }

                if (best == null)
                {
                    var dump = new StringBuilder();
                    foreach (var d in devs) dump.Append($"[{DriverSuffix(d.Type) ?? d.Type.ToString()}:{d.Name}] ");
                    Log($"RestoreIoDevice {slot}: no Voicemeeter {(isInput ? "input" : "output")} device matches '{targetName}'. enumerated: {dump}");
                    return false;
                }

                var m = best.Value;
                string drv = DriverSuffix(m.Type)!;
                int rc = VBVMR_SetParameterStringW($"{slot}.device.{drv}", m.Name);
                if (rc < 0)
                {
                    Log($"RestoreIoDevice {slot}: set .{drv} '{m.Name}' rc={rc} (rejected, give up)");
                    return false;
                }
                // Don't poll device.name here: the swap is async (lags >1s) and this
                // lock is shared with the 50ms UI peak meter. rc>=0 + enumerated match
                // means it's applied; the caller debounces re-sets during the lag.
                Log($"RestoreIoDevice {slot}: drift '{current}' -> set .{drv} '{m.Name}' rc={rc} (applied)");
                return true;
            }
            catch (Exception ex) { Log($"RestoreIoDevice {slot} exception: {ex.Message}"); return false; }
        }
    }

    private static string DriverLabel(int type) => type switch
    {
        1 => "MME",
        3 => "WDM",
        4 => "KS",
        5 => "ASIO",
        _ => "?",
    };

    // The devices Voicemeeter offers for a Strip (input) / Bus (output), exactly as it
    // enumerates them — used to populate the click-to-select menu.
    public static List<VoicemeeterDevice> GetSelectableDevices(bool isInput)
    {
        lock (_gate)
        {
            if (!EnsureLoggedIn()) return [];
            try
            {
                // The user just opened the picker — they may have (un)plugged a device, so
                // refresh fresh here and drop the cache so the next poll's row agrees.
                InvalidateSysSnapshot();
                var sys = EnumerateSystemEndpoints();
                var modelRank = BuildModelRank(sys);
                // Endpoints already claimed, tracked per driver type: within one driver
                // the order pairing disambiguates [A]/[B], but the same device legitimately
                // recurs across drivers (WDM "Headphones" and KS "Headphones") and must be
                // free to match again there.
                var consumedByType = new Dictionary<int, HashSet<int>>();
                var result = new List<VoicemeeterDevice>();
                foreach (var d in EnumVmDevices(isInput))
                {
                    var suf = DriverSuffix(d.Type);
                    if (suf == null) continue;
                    if (!consumedByType.TryGetValue(d.Type, out var consumed))
                        consumedByType[d.Type] = consumed = new HashSet<int>();
                    var display = ResolveDisplayName(d.Name, d.Hwid, isInput, sys, consumed, modelRank);
                    result.Add(new VoicemeeterDevice(d.Type, suf, DriverLabel(d.Type), d.Name, display));
                }
                // Log only the enriched mappings (+ a count) rather than the whole list —
                // enough to diagnose name resolution without flooding the log each open.
                var enriched = result.Where(r => r.Name != r.DisplayName)
                    .Select(r => $"[{r.DriverLabel}:{r.Name}=>{r.DisplayName}]").ToList();
                Log($"GetSelectableDevices({(isInput ? "in" : "out")}): {result.Count} devices"
                    + (enriched.Count > 0 ? ", enriched " + string.Join(" ", enriched) : ""));
                return result;
            }
            catch (Exception ex) { Log($"GetSelectableDevices exception: {ex.Message}"); return []; }
        }
    }

    // User picked a device from the menu: set it on the chosen driver.
    public static bool SetIoDevice(bool isInput, int index, int driverType, string name)
    {
        if (string.IsNullOrEmpty(name)) return false;
        var suf = DriverSuffix(driverType);
        if (suf == null) return false;
        lock (_gate)
        {
            if (!EnsureLoggedIn()) return false;
            string slot = isInput ? $"Strip[{index}]" : $"Bus[{index}]";
            try
            {
                int rc = VBVMR_SetParameterStringW($"{slot}.device.{suf}", name);
                // rc>=0 means Voicemeeter accepted it. The endpoint swap is async and
                // device.name lags well past a second — do NOT sleep/poll here: this
                // lock is shared with the 50ms UI peak meter, so a long hold freezes
                // the UI. One immediate read is just for the diagnostic log.
                if (rc < 0)
                {
                    Log($"SetIoDevice {slot}.device.{suf} = '{name}' rc={rc} (rejected)");
                    return false;
                }
                Log($"SetIoDevice {slot}.device.{suf} = '{name}' rc={rc} accepted "
                    + $"(name now '{ReadStringParam($"{slot}.device.name")}', swap is async)");
                return true;
            }
            catch (Exception ex) { Log($"SetIoDevice {slot} exception: {ex.Message}"); return false; }
        }
    }

    private static bool NamesMatch(string? a, string? b)
    {
        if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return false;
        if (a.Equals(b, StringComparison.OrdinalIgnoreCase)) return true;
        // Prefix match handles MME's ~31-char truncation, but a " [A]"/" [B]" suffix is
        // Voicemeeter's disambiguator for DIFFERENT same-named devices (e.g. "Headset" vs
        // "Headset [A]") — treating those as equal restores/identifies the wrong device.
        return IsTruncationPrefix(a, b) || IsTruncationPrefix(b, a);
    }

    // True if `longer` starts with `shorter` and the extra text is a genuine continuation
    // (an MME truncation), not a " [X]" disambiguator marking a different device.
    private static bool IsTruncationPrefix(string longer, string shorter)
    {
        if (!longer.StartsWith(shorter, StringComparison.OrdinalIgnoreCase)) return false;
        var extra = longer.Substring(shorter.Length).TrimStart();
        return !extra.StartsWith("[", StringComparison.Ordinal);
    }

    // VM-device-name -> enriched display name, keyed "O|<name>" (output/Bus) or "I|<name>"
    // (input/Strip). Rebuilt by GetIoState (off the UI thread, under _gate) and read
    // lock-free by the UI via DisplayNameFor — the reference is swapped atomically and the
    // map itself is never mutated after publishing. Only holds entries whose display
    // differs from the raw name (BT "Headphones [A]", truncated MME, generic KS).
    private static volatile Dictionary<string, string> _displayMap = new(StringComparer.OrdinalIgnoreCase);

    // Enumerating Windows endpoints (with per-device property reads) and rebuilding the
    // display map is the heavy part of GetIoState, which runs as often as every 750ms while
    // a lock is enforced. The device set rarely changes, so cache the snapshot for a few
    // seconds; the per-poll device.name reads that drive drift detection stay fresh.
    private sealed record SysSnapshot(List<SysEndpoint> Endpoints, List<string> Names, Dictionary<string, string> DisplayMap);
    private static SysSnapshot? _sysSnapshot;
    private static DateTime _sysSnapshotAt;
    private static readonly TimeSpan SysSnapshotTtl = TimeSpan.FromSeconds(4);

    // Caller must hold _gate.
    private static SysSnapshot GetSysSnapshot()
    {
        if (_sysSnapshot != null && DateTime.UtcNow - _sysSnapshotAt < SysSnapshotTtl)
            return _sysSnapshot;
        var eps = EnumerateSystemEndpoints();
        var snap = new SysSnapshot(eps, eps.Select(e => e.Friendly).ToList(), BuildDisplayMap(eps));
        _sysSnapshot = snap;
        _sysSnapshotAt = DateTime.UtcNow;
        _displayMap = snap.DisplayMap;
        return snap;
    }

    // Drop the cached snapshot so the next GetSysSnapshot rebuilds fresh — used when the
    // user opens the picker (they may have just (un)plugged a device) and on session reset.
    private static void InvalidateSysSnapshot() => _sysSnapshot = null;

    private static Dictionary<string, string> BuildDisplayMap(List<SysEndpoint> sys)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var modelRank = BuildModelRank(sys);
        foreach (var isInput in new[] { false, true })
        {
            // Same per-driver order pairing as GetSelectableDevices, so the current-device
            // row shows exactly what the picker offers.
            var consumedByType = new Dictionary<int, HashSet<int>>();
            foreach (var d in EnumVmDevices(isInput))
            {
                if (DriverSuffix(d.Type) == null) continue;
                if (!consumedByType.TryGetValue(d.Type, out var consumed))
                    consumedByType[d.Type] = consumed = new HashSet<int>();
                var display = ResolveDisplayName(d.Name, d.Hwid, isInput, sys, consumed, modelRank);
                if (!string.Equals(display, d.Name, StringComparison.Ordinal))
                    map[$"{(isInput ? "I" : "O")}|{d.Name}"] = display;
            }
        }
        return map;
    }

    // Enriched display name for a Voicemeeter device name (the value of Bus/Strip
    // device.name). Returns the input unchanged when there's no enrichment. Lock-free.
    public static string? DisplayNameFor(string? vmName, bool isInput)
    {
        if (string.IsNullOrEmpty(vmName)) return vmName;
        return _displayMap.TryGetValue($"{(isInput ? "I" : "O")}|{vmName}", out var d) ? d : vmName;
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

    // A Windows audio endpoint paired with the hardware id Voicemeeter also reports for
    // it, so we can map a Voicemeeter device back to the friendly name Windows shows.
    // State lets one enumeration serve both jobs: Active-only feeds the "device missing"
    // (red) check, while the full set incl. Unplugged feeds name enrichment.
    private record SysEndpoint(string Friendly, string Hwid, DataFlow Flow, DeviceState State);

    // The property whose value is the exact string VBVMR_*_GetDeviceDescW returns as the
    // device hwid (e.g. "USB\VID_0DB0&PID_422D&MI_00", "HDAUDIO\FUNC_01&VEN_10DE&DEV_00A4",
    // "BTHENUM\{0000110b-...}"). For Bluetooth, Windows reports a longer instance id that
    // *starts with* Voicemeeter's value, so matching is prefix-tolerant.
    private static readonly Guid PkeyHardwareId = new("a8b865dd-2e3d-4094-ad97-e593a70c75d6");
    private const int PkeyHardwareIdPid = 8;

    private static List<SysEndpoint> EnumerateSystemEndpoints()
    {
        var list = new List<SysEndpoint>();
        try
        {
            using var enumerator = new MMDeviceEnumerator();
            // Include Unplugged/Disabled, not just Active: a paired-but-disconnected
            // Bluetooth headphone is Unplugged, yet its friendly name ("Headphones (HD
            // 450BT)") is what we want to show. Skip NotPresent to avoid stale ghosts.
            const DeviceState states = DeviceState.Active | DeviceState.Unplugged | DeviceState.Disabled;
            foreach (var d in enumerator.EnumerateAudioEndPoints(DataFlow.All, states))
            {
                try
                {
                    var friendly = d.FriendlyName;
                    string? hwid = null;
                    var store = d.Properties;
                    for (int i = 0; i < store.Count; i++)
                    {
                        var p = store[i];
                        if (p.Key.propertyId == PkeyHardwareIdPid && p.Key.formatId == PkeyHardwareId)
                        {
                            try { hwid = p.Value?.ToString(); } catch { }
                            break;
                        }
                    }
                    // Keep endpoints even when the hwid property is absent: they still
                    // count for the "present"/not-missing check (so a device without it
                    // doesn't falsely show red). Enrichment just skips a no-hwid endpoint,
                    // since HwidRelated requires a hwid on both sides.
                    if (!string.IsNullOrEmpty(friendly))
                        list.Add(new SysEndpoint(friendly, hwid ?? "", d.DataFlow, d.State));
                }
                catch { }
                finally { d.Dispose(); }
            }
        }
        catch { }
        return list;
    }

    // Canonical device order, taken from the RENDER endpoints (whose enumeration order
    // matches Voicemeeter's device order — validated), keyed by the model part of the
    // friendly name ("Headphones (HD 450BT)" -> "HD 450BT"). Used to order matches so
    // pairing follows Voicemeeter even when an endpoint's own flow enumerates in a
    // different order (Windows lists the two Bluetooth headphones HD-450BT-first under
    // Render but BX17-first under Capture).
    private static Dictionary<string, int> BuildModelRank(List<SysEndpoint> sys)
    {
        var rank = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var e in sys)
            if (e.Flow == DataFlow.Render)
            {
                var m = ModelPart(e.Friendly);
                if (!rank.ContainsKey(m)) rank[m] = rank.Count;
            }
        return rank;
    }

    private static string ModelPart(string friendly)
    {
        int p = friendly.IndexOf(" (", StringComparison.Ordinal);
        if (p < 0) return friendly;
        var s = friendly.Substring(p + 2);
        return s.EndsWith(')') ? s[..^1] : s;
    }

    // Voicemeeter's KS names are generic ("Headphones") and its MME names are truncated
    // to ~31 chars, so neither identifies the actual model. The Windows friendly name
    // carries it ("Headphones (HD 450BT)"). Match by hardware id + name affinity and
    // return the friendly name.
    //
    // The hard case is two same-profile Bluetooth headphones: identical hwid AND jack name,
    // no model/MAC over the Remote API — Voicemeeter disambiguates them only as
    // "Headphones [A]"/"[B]" by ENUMERATION ORDER. We mirror that: callers pass VM devices
    // in Voicemeeter's order with a per-driver `consumed` set, and each device claims the
    // best endpoint not yet taken. Ordering key, lowest wins: (1) same data flow preferred —
    // but HFP "Headset" output has no Render endpoint, so fall back to its Capture endpoint;
    // (2) higher name affinity; (3) canonical model rank, so the pairing matches Voicemeeter
    // even when that endpoint's flow enumerates in a different order; (4) raw index. The
    // `consumed` set is reset per driver type so one physical device still resolves under
    // both its WDM and KS entries.
    private static string ResolveDisplayName(string vmName, string vmHwid, bool isInput,
        List<SysEndpoint> sys, HashSet<int> consumed, Dictionary<string, int> modelRank)
    {
        if (string.IsNullOrEmpty(vmHwid)) return vmName;
        var flow = isInput ? DataFlow.Capture : DataFlow.Render;
        int bestIdx = -1;
        (int, int, int, int) bestKey = default;
        for (int i = 0; i < sys.Count; i++)
        {
            if (consumed.Contains(i)) continue;
            var e = sys[i];
            if (!HwidRelated(e.Hwid, vmHwid)) continue;
            int aff = NameAffinity(e.Friendly, vmName);
            if (aff == 0) continue;
            int rank = modelRank.TryGetValue(ModelPart(e.Friendly), out var r) ? r : int.MaxValue;
            var key = (e.Flow == flow ? 0 : 1, -aff, rank, i);
            if (bestIdx < 0 || key.CompareTo(bestKey) < 0) { bestKey = key; bestIdx = i; }
        }
        if (bestIdx < 0) return vmName;  // offline / ASIO / no match
        consumed.Add(bestIdx);
        var friendly = sys[bestIdx].Friendly;
        return friendly.Equals(vmName, StringComparison.OrdinalIgnoreCase) ? vmName : friendly;
    }

    private static bool HwidRelated(string a, string b)
    {
        if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return false;
        return a.Equals(b, StringComparison.OrdinalIgnoreCase)
            || a.StartsWith(b, StringComparison.OrdinalIgnoreCase)
            || b.StartsWith(a, StringComparison.OrdinalIgnoreCase);
    }

    // How strongly a Windows friendly name corresponds to a Voicemeeter device name.
    // Higher = more confident; 0 = unrelated (caller keeps the Voicemeeter name).
    private static int NameAffinity(string friendly, string vmName)
    {
        if (friendly.Equals(vmName, StringComparison.OrdinalIgnoreCase)) return 100;
        // MME truncates to ~31 chars: the Voicemeeter name is a prefix of the full one.
        if (friendly.StartsWith(vmName, StringComparison.OrdinalIgnoreCase)) return 90;
        // Windows friendly name = "<jack> (<device>)"; KS gives just the jack ("Headphones").
        var jack = JackName(friendly);
        if (jack.Equals(vmName, StringComparison.OrdinalIgnoreCase)) return 80;
        // KS may add a disambiguator: "Headphones [A]" vs jack "Headphones".
        if (jack.Length >= 3 && vmName.StartsWith(jack, StringComparison.OrdinalIgnoreCase)) return 70;
        return 0;
    }

    private static string JackName(string friendly)
    {
        int p = friendly.IndexOf(" (", StringComparison.Ordinal);
        return p > 0 ? friendly.Substring(0, p) : friendly;
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

    // Same marshaling split as the getter above: ASCII param name, Unicode value
    // (device friendly names can contain CJK characters).
    [DllImport(DllName, CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_SetParameterStringW(
        [MarshalAs(UnmanagedType.LPStr)] string param,
        [MarshalAs(UnmanagedType.LPWStr)] string value);

    // Device enumeration — the only reliable source for the name + driver type that
    // VBVMR_SetParameterString expects when selecting a Strip/Bus hardware device.
    [DllImport(DllName, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_Output_GetDeviceNumber();

    [DllImport(DllName, CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_Output_GetDeviceDescW(
        int zindex, out int nType, StringBuilder name, StringBuilder hardwareId);

    [DllImport(DllName, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_Input_GetDeviceNumber();

    [DllImport(DllName, CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
    private static extern int VBVMR_Input_GetDeviceDescW(
        int zindex, out int nType, StringBuilder name, StringBuilder hardwareId);
}
