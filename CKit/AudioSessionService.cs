using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Media;
using NAudio.CoreAudioApi;

namespace AudioDeviceSwitcher;

public record AppAudioSessionInfo(
    uint ProcessId,
    string DisplayName,
    string? ExecutablePath,
    ImageSource? Icon,
    IReadOnlyList<string>? SessionInstanceKeys,
    float Volume,
    bool Muted);

public static class AudioSessionService
{
    private const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern nint OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(nint hObject);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool QueryFullProcessImageName(nint hProcess, uint dwFlags, StringBuilder lpExeName, ref uint lpdwSize);

    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
    private static extern int SHLoadIndirectString(string pszSource, StringBuilder pszOutBuf, int cchOutBuf, nint ppvReserved);

    // Process.MainModule needs PROCESS_VM_READ and fails with access-denied for processes
    // running at a higher integrity level (e.g. elevated apps) even though their audio
    // session is visible to us. QueryFullProcessImageName only needs the "limited info"
    // access right, which Windows grants regardless of integrity level (same trick Task
    // Manager uses to show exe paths for elevated processes).
    private static string? GetProcessImagePath(uint pid)
    {
        var handle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
        if (handle == 0) return null;
        try
        {
            var sb = new StringBuilder(1024);
            uint size = (uint)sb.Capacity;
            return QueryFullProcessImageName(handle, 0, sb, ref size) ? sb.ToString(0, (int)size) : null;
        }
        finally { CloseHandle(handle); }
    }

    // Some hosts (observed with BlueStacks' HD-Player.exe) harden their process DACL enough
    // that even OpenProcess with PROCESS_QUERY_LIMITED_INFORMATION is denied to a non-admin
    // caller. WMI's provider host resolves Win32_Process fields through a different code
    // path and succeeds in that case where direct OpenProcess-based calls above don't.
    // Looks up ExecutablePath/CommandLine for every requested PID in one round trip — each
    // WMI query costs tens of milliseconds of provider-host overhead, and windowless
    // processes (the common case for audio-only sessions) all need this fallback, so calling
    // it once per PID was the dominant cost of opening the app-audio window.
    private static Dictionary<uint, (string? ExePath, string? CommandLine)> GetProcessWmiInfoBatch(IReadOnlyCollection<uint> pids)
    {
        var result = new Dictionary<uint, (string?, string?)>();
        if (pids.Count == 0) return result;
        try
        {
            var clause = string.Join(" OR ", pids.Select(id => $"ProcessId = {id}"));
            using var searcher = new ManagementObjectSearcher(
                $"SELECT ProcessId, ExecutablePath, CommandLine FROM Win32_Process WHERE {clause}");
            using var results = searcher.Get();
            foreach (ManagementObject mo in results)
            {
                var pid = Convert.ToUInt32(mo["ProcessId"]);
                result[pid] = (mo["ExecutablePath"] as string, mo["CommandLine"] as string);
            }
        }
        catch { }
        return result;
    }

    // BlueStacks' multi-instance launcher shortcuts pass "--instance <name>" on the command
    // line so each shortcut can be told apart even when several instances share the exact
    // same exe and process name, and the audio-owning process itself has no visible window
    // to read a distinguishing title from.
    private static string? ExtractInstanceHint(string? commandLine)
    {
        if (string.IsNullOrEmpty(commandLine)) return null;
        var match = Regex.Match(commandLine, @"--instance[=\s]+""?([^""\s]+)""?");
        return match.Success ? match.Groups[1].Value : null;
    }

    // Mirrors the priority Windows' own Volume Mixer uses: the session's own display name
    // (set explicitly by the app, rarely used) beats the exe's FileDescription version
    // resource (e.g. "Google Chrome" instead of the raw "chrome"), which beats the bare
    // process name.
    private static string ResolveDisplayName(uint pid, string processName, string? exePath, string? sessionDisplayName)
    {
        if (!string.IsNullOrEmpty(sessionDisplayName)) return sessionDisplayName;

        if (!string.IsNullOrEmpty(exePath))
        {
            try
            {
                var desc = FileVersionInfo.GetVersionInfo(exePath).FileDescription;
                if (!string.IsNullOrWhiteSpace(desc)) return desc!.Trim();
            }
            catch { }
        }

        if (!string.IsNullOrEmpty(processName)) return processName;
        return $"PID {pid}";
    }

    // Some sessions (mostly packaged/UWP apps) report an indirect string reference like
    // "@{Microsoft.Something_8wekyb3d8bbwe?ms-resource://.../DisplayName}" instead of plain
    // text; resolve it the same way Explorer/Volume Mixer do.
    private static string? ResolveIndirectString(string source)
    {
        if (!source.StartsWith('@')) return source;
        try
        {
            var sb = new StringBuilder(1024);
            var hr = SHLoadIndirectString(source, sb, sb.Capacity, 0);
            return hr == 0 ? sb.ToString() : null;
        }
        catch { return null; }
    }

    private readonly record struct RawSession(
        uint Pid, string? ExplicitName, string InstanceKey,
        string ProcessName, string? ExePath, string? WindowTitle, string? InstanceHint,
        float Volume, bool Muted);

    // Walks every session on every active render/capture device and returns one raw record
    // per session (no grouping yet — a PID commonly owns several sessions, e.g. separate
    // render/capture streams). InstanceKey is IAudioSessionControl2's own per-session
    // identifier: stable for the life of that COM session object, so it's the one thing here
    // that's safe to hold onto and match against later (unlike DisplayName, which some apps
    // mutate after creation, e.g. to show now-playing info).
    private static List<RawSession> EnumerateRawSessions()
    {
        // Pid/explicit-name/instance-key are cheap to collect while walking the COM session
        // tree; the WMI fallback (needed for windowless processes) is deferred to a single
        // batched query after this loop instead of one query per PID — see
        // GetProcessWmiInfoBatch.
        var pending = new List<(uint Pid, string? ExplicitName, string InstanceKey, float Volume, bool Muted)>();
        var procCache = new Dictionary<uint, (string ProcessName, string? ExePath, string? WindowTitle)>();
        var needsWmi = new HashSet<uint>();
        using var enumerator = new MMDeviceEnumerator();

        foreach (var flow in new[] { DataFlow.Render, DataFlow.Capture })
        {
            MMDeviceCollection devices;
            try { devices = enumerator.EnumerateAudioEndPoints(flow, DeviceState.Active); }
            catch (COMException) { continue; }

            foreach (var device in devices)
            {
                try
                {
                    AudioSessionManager? manager;
                    try { manager = device.AudioSessionManager; }
                    catch (COMException) { continue; }

                    try
                    {
                        var sessions = manager.Sessions;
                        if (sessions == null) continue;

                        for (int i = 0; i < sessions.Count; i++)
                        {
                            var session = sessions[i];
                            try
                            {
                                uint pid;
                                try { pid = session.GetProcessID; }
                                catch { continue; }
                                if (pid == 0) continue;

                                if (!procCache.TryGetValue(pid, out var procInfo))
                                {
                                    try
                                    {
                                        using var proc = Process.GetProcessById((int)pid);
                                        var processName = proc.ProcessName;
                                        string? exePath = null;
                                        try { exePath = proc.MainModule?.FileName; } catch { }
                                        if (string.IsNullOrEmpty(exePath)) exePath = GetProcessImagePath(pid);
                                        string? windowTitle = null;
                                        try { windowTitle = proc.MainWindowTitle; } catch { }

                                        if (string.IsNullOrEmpty(exePath) || string.IsNullOrWhiteSpace(windowTitle))
                                            needsWmi.Add(pid);

                                        procInfo = (processName, exePath, windowTitle);
                                    }
                                    catch
                                    {
                                        continue; // process gone
                                    }
                                    procCache[pid] = procInfo;
                                }

                                string? rawName = null;
                                try { rawName = session.DisplayName; } catch { }
                                var explicitName = string.IsNullOrEmpty(rawName) ? null : ResolveIndirectString(rawName);

                                string? instanceKey = null;
                                try { instanceKey = session.GetSessionInstanceIdentifier; } catch { }
                                if (string.IsNullOrEmpty(instanceKey)) instanceKey = $"{pid}#{flow}#{i}#{Guid.NewGuid()}";

                                // Captured here (the session object is already open) instead
                                // of via a second GetAppVolume call later — that call would
                                // re-walk every device/session on every device from scratch,
                                // effectively repeating this whole enumeration once per row.
                                float volume = 1f;
                                bool muted = false;
                                try
                                {
                                    var sav = session.SimpleAudioVolume;
                                    volume = sav.Volume;
                                    muted = sav.Mute;
                                }
                                catch { }

                                pending.Add((pid, explicitName, instanceKey, volume, muted));
                            }
                            finally { try { session.Dispose(); } catch { } }
                        }
                    }
                    finally { try { manager.Dispose(); } catch { } }
                }
                finally { device.Dispose(); }
            }
        }

        var wmiInfo = GetProcessWmiInfoBatch(needsWmi);
        var instanceHints = new Dictionary<uint, string?>();
        foreach (var pid in needsWmi)
        {
            var info = procCache[pid];
            if (!wmiInfo.TryGetValue(pid, out var wmi)) continue;
            var exePath = string.IsNullOrEmpty(info.ExePath) ? wmi.ExePath : info.ExePath;
            procCache[pid] = (info.ProcessName, exePath, info.WindowTitle);
            instanceHints[pid] = ExtractInstanceHint(wmi.CommandLine);
        }

        var raw = new List<RawSession>(pending.Count);
        foreach (var p in pending)
        {
            var info = procCache[p.Pid];
            instanceHints.TryGetValue(p.Pid, out var instanceHint);
            raw.Add(new RawSession(p.Pid, p.ExplicitName, p.InstanceKey,
                info.ProcessName, info.ExePath, info.WindowTitle, instanceHint, p.Volume, p.Muted));
        }

        return raw;
    }

    // Enumerates active audio sessions across all render + capture devices, one row per
    // logical app instance. Grouping key is (PID, explicit session DisplayName): a process
    // usually owns several sessions (e.g. a separate render and capture stream) that all
    // belong to the same logical app and should collapse into one row — the old, PID-only
    // behavior. But a multi-instance host process (e.g. an emulator running several renamed
    // "channels" through one shared process) tags each instance's sessions with its own
    // explicit DisplayName; when that's present it identifies a distinct instance, so those
    // sessions must NOT collapse with the rest of that PID's sessions.
    // Skips the System sounds session (PID 0) to keep the UI focused on user apps.
    public static List<AppAudioSessionInfo> GetActiveAppSessions()
    {
        var groups = EnumerateRawSessions()
            .GroupBy(r => (r.Pid, Key: r.ExplicitName?.ToLowerInvariant() ?? ""));

        var list = new List<AppAudioSessionInfo>();
        foreach (var group in groups)
        {
            var first = group.First();
            var hasExplicitName = !string.IsNullOrEmpty(first.ExplicitName);

            string displayName = ResolveDisplayName(first.Pid, first.ProcessName, first.ExePath, first.ExplicitName);

            // A session that never set its own DisplayName resolved to the generic
            // FileDescription/process name above, the same for every instance of a
            // multi-instance app — prefer the window title if it carries more specific
            // information, since that's the other place a user-assigned instance name (e.g.
            // a renamed emulator "channel") usually surfaces. Some hosts (e.g. BlueStacks'
            // background audio-owning process) have no window at all to read a title from;
            // fall back to the --instance launch-arg hint in that case.
            if (!hasExplicitName)
            {
                var extra = !string.IsNullOrWhiteSpace(first.WindowTitle) &&
                    !string.Equals(first.WindowTitle, displayName, StringComparison.CurrentCultureIgnoreCase)
                    ? first.WindowTitle
                    : !string.IsNullOrWhiteSpace(first.InstanceHint) &&
                      !string.Equals(first.InstanceHint, displayName, StringComparison.CurrentCultureIgnoreCase)
                        ? first.InstanceHint
                        : null;
                if (extra != null) displayName = extra;
            }

            ImageSource? icon = null;
            if (!string.IsNullOrEmpty(first.ExePath))
            {
                try { icon = DeviceIconHelper.GetExeIcon(first.ExePath); } catch { }
            }

            // Only carry precise session keys for explicitly-named groups (distinct
            // instances sharing a PID). Unnamed sessions already get PID-wide control via a
            // null key, which is exactly right for them and doesn't need this precision.
            var sessionKeys = hasExplicitName ? group.Select(r => r.InstanceKey).ToList() : null;

            list.Add(new AppAudioSessionInfo(first.Pid, displayName, first.ExePath, icon, sessionKeys,
                first.Volume, first.Muted));
        }

        list.Sort((a, b) => string.Compare(a.DisplayName, b.DisplayName, StringComparison.CurrentCultureIgnoreCase));
        return list;
    }

    // Per-application volume/mute via the audio session's ISimpleAudioVolume. A process may
    // own several sessions (e.g. separate render/capture streams, or several logical
    // "channels" sharing one host process — see GetActiveAppSessions). sessionKeys, when
    // given, are the exact IAudioSessionControl2 instance identifiers captured for this row;
    // only those sessions are affected. When null, every session owned by the PID is
    // affected — the correct behavior for ordinary apps that never set a DisplayName.
    //
    // If sessionKeys is given but none of them are live anymore (the underlying stream was
    // torn down and recreated between enumeration and this call), fall back to every session
    // on the PID rather than silently doing nothing: a slightly-too-broad effect beats a
    // control that appears to just not work.
    private static void ForEachSession(uint pid, IReadOnlyList<string>? sessionKeys, Action<AudioSessionControl> action)
    {
        int matched = 0;
        ForEachSessionCore(pid, sessionKeys, s => { matched++; action(s); });
        if (matched == 0 && sessionKeys != null)
            ForEachSessionCore(pid, null, action);
    }

    private static void ForEachSessionCore(uint pid, IReadOnlyList<string>? sessionKeys, Action<AudioSessionControl> action)
    {
        using var enumerator = new MMDeviceEnumerator();
        foreach (var flow in new[] { DataFlow.Render, DataFlow.Capture })
        {
            MMDeviceCollection devices;
            try { devices = enumerator.EnumerateAudioEndPoints(flow, DeviceState.Active); }
            catch (COMException) { continue; }

            foreach (var device in devices)
            {
                try
                {
                    AudioSessionManager? manager;
                    try { manager = device.AudioSessionManager; }
                    catch (COMException) { continue; }

                    try
                    {
                        var sessions = manager.Sessions;
                        if (sessions == null) continue;
                        for (int i = 0; i < sessions.Count; i++)
                        {
                            var session = sessions[i];
                            try
                            {
                                uint spid;
                                try { spid = session.GetProcessID; }
                                catch { continue; }
                                if (spid != pid) continue;

                                if (sessionKeys != null)
                                {
                                    string? instanceKey = null;
                                    try { instanceKey = session.GetSessionInstanceIdentifier; } catch { }
                                    if (instanceKey == null || !sessionKeys.Contains(instanceKey, StringComparer.Ordinal))
                                        continue;
                                }

                                try { action(session); } catch { }
                            }
                            finally { try { session.Dispose(); } catch { } }
                        }
                    }
                    finally { try { manager.Dispose(); } catch { } }
                }
                finally { device.Dispose(); }
            }
        }
    }

    public static void SetAppVolume(uint pid, float volume, IReadOnlyList<string>? sessionKeys = null)
    {
        ForEachSession(pid, sessionKeys, s => s.SimpleAudioVolume.Volume = volume);
    }

    public static void SetAppMute(uint pid, bool mute, IReadOnlyList<string>? sessionKeys = null)
    {
        ForEachSession(pid, sessionKeys, s => s.SimpleAudioVolume.Mute = mute);
    }
}
