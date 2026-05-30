using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Media;
using NAudio.CoreAudioApi;

namespace AudioDeviceSwitcher;

public record AppAudioSessionInfo(
    uint ProcessId,
    string DisplayName,
    string? ExecutablePath,
    ImageSource? Icon);

public static class AudioSessionService
{
    // Enumerates active audio sessions across all render + capture devices, deduped by PID.
    // Skips the System sounds session (PID 0) to keep the UI focused on user apps.
    public static List<AppAudioSessionInfo> GetActiveAppSessions()
    {
        var seen = new Dictionary<uint, AppAudioSessionInfo>();
        using var enumerator = new MMDeviceEnumerator();

        foreach (var flow in new[] { DataFlow.Render, DataFlow.Capture })
        {
            MMDeviceCollection devices;
            try
            {
                devices = enumerator.EnumerateAudioEndPoints(flow, DeviceState.Active);
            }
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
                                if (seen.ContainsKey(pid)) continue;

                                string displayName = "";
                                string? exePath = null;
                                ImageSource? icon = null;

                                try
                                {
                                    using var proc = Process.GetProcessById((int)pid);
                                    displayName = proc.ProcessName;
                                    try { exePath = proc.MainModule?.FileName; } catch { }
                                }
                                catch
                                {
                                    continue; // process gone
                                }

                                if (string.IsNullOrEmpty(displayName))
                                {
                                    try { displayName = session.DisplayName ?? ""; } catch { }
                                }
                                if (string.IsNullOrEmpty(displayName)) displayName = $"PID {pid}";

                                if (!string.IsNullOrEmpty(exePath))
                                {
                                    try { icon = DeviceIconHelper.GetExeIcon(exePath); } catch { }
                                }

                                seen[pid] = new AppAudioSessionInfo(pid, displayName, exePath, icon);
                            }
                            finally { try { session.Dispose(); } catch { } }
                        }
                    }
                    finally { try { manager.Dispose(); } catch { } }
                }
                finally { device.Dispose(); }
            }
        }

        var list = seen.Values.ToList();
        list.Sort((a, b) => string.Compare(a.DisplayName, b.DisplayName, StringComparison.CurrentCultureIgnoreCase));
        return list;
    }

    // Per-application volume/mute via the audio session's ISimpleAudioVolume. A process may own
    // sessions on several devices; reads return the first match, writes apply to every match.
    private static void ForEachSession(uint pid, Action<AudioSessionControl> action)
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

    public static (float Volume, bool Muted)? GetAppVolume(uint pid)
    {
        (float, bool)? result = null;
        ForEachSession(pid, s =>
        {
            if (result != null) return;
            var v = s.SimpleAudioVolume;
            result = (v.Volume, v.Mute);
        });
        return result;
    }

    public static void SetAppVolume(uint pid, float volume)
    {
        ForEachSession(pid, s => s.SimpleAudioVolume.Volume = volume);
    }

    public static void SetAppMute(uint pid, bool mute)
    {
        ForEachSession(pid, s => s.SimpleAudioVolume.Mute = mute);
    }
}
