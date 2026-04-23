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
                AudioSessionManager manager;
                try
                {
                    manager = device.AudioSessionManager;
                }
                catch (COMException) { continue; }

                var sessions = manager.Sessions;
                if (sessions == null) continue;

                for (int i = 0; i < sessions.Count; i++)
                {
                    var session = sessions[i];
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
            }
        }

        var list = seen.Values.ToList();
        list.Sort((a, b) => string.Compare(a.DisplayName, b.DisplayName, StringComparison.CurrentCultureIgnoreCase));
        return list;
    }
}
