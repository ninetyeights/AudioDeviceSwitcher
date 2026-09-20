using System.Diagnostics;
using System.IO;
using NAudio.CoreAudioApi;

namespace AudioDeviceSwitcher;

public record ProfileApplyResult(
    int AppliedOverrideCount,
    List<string> SkippedAppNames,
    List<string> MissingDeviceNames,
    VoicemeeterRestartStatus VoicemeeterStatus);

public static class ProfileApplyService
{
    public static event Action<DeviceProfile>? Applying;
    private static Guid? _selectedProfileId;
    public static Guid? SelectedProfileId => _selectedProfileId;

    public static DeviceProfile? FindActiveProfile(string? playback, string? recording)
    {
        var profiles = ProfileService.GetAll();
        bool Matches(DeviceProfile p) => p.PlaybackDeviceId == playback && p.RecordingDeviceId == recording;
        return profiles.Find(p => p.Id == SettingsService.Load().LockedProfileId && Matches(p))
            ?? profiles.Find(p => p.Id == _selectedProfileId && Matches(p))
            ?? profiles.Find(Matches);
    }

    public static ProfileApplyResult Apply(DeviceProfile profile)
    {
        _selectedProfileId = profile.Id;
        Applying?.Invoke(profile);
        var allPlayback = AudioDeviceService.GetPlaybackDevices();
        var allRecording = AudioDeviceService.GetRecordingDevices();

        // Skip devices not in the active set (e.g. bluetooth headphones disconnected).
        // COM SetDefaultEndpoint silently no-ops for such ids — without this guard the
        // user sees a "switched" toast while nothing actually changed.
        bool playbackAvailable = profile.PlaybackDeviceId != null
            && allPlayback.Any(d => string.Equals(d.Id, profile.PlaybackDeviceId, StringComparison.Ordinal));
        bool recordingAvailable = profile.RecordingDeviceId != null
            && allRecording.Any(d => string.Equals(d.Id, profile.RecordingDeviceId, StringComparison.Ordinal));

        var missing = new List<string>();
        if (profile.PlaybackDeviceId != null && !playbackAvailable)
            missing.Add($"播放: {profile.PlaybackDeviceName ?? "未知"}");
        if (profile.RecordingDeviceId != null && !recordingAvailable)
            missing.Add($"录音: {profile.RecordingDeviceName ?? "未知"}");

        // Clear persisted preferences, including apps that are not running. A PID-based
        // reset misses those apps and they reopen on their previous profile's devices.
        // Do this even when system defaults already match and no overrides exist.
        // Let failure propagate: applying overrides on a stale baseline is not success.
        AppAudioRoutingService.ClearAll();

        if (playbackAvailable)
            AudioDeviceService.SetDefaultDevice(profile.PlaybackDeviceId!);
        if (recordingAvailable)
            AudioDeviceService.SetDefaultDevice(profile.RecordingDeviceId!);

        // The profile's explicit overrides are the only exceptions to follow-system.
        var thisProfileExes = profile.AppOverrides
            .Select(o => o.ExePath)
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var runningByPath = GetRunningProcessesByPath(thisProfileExes);

        // Merge audio-session PIDs: Chrome-style apps often spawn child processes whose
        // session PID differs from the main-exe PID. Setting only on one leaves the other
        // unset — per-app override works (route changes) but queries via the other PID
        // return "follow system", confusing the UI.
        foreach (var s in thisProfileExes.Count > 0 ? AudioSessionService.GetActiveAppSessions() : [])
        {
            if (string.IsNullOrEmpty(s.ExecutablePath)) continue;
            if (!thisProfileExes.Contains(s.ExecutablePath)) continue;
            if (!runningByPath.TryGetValue(s.ExecutablePath, out var list))
                runningByPath[s.ExecutablePath] = list = new List<uint>();
            if (!list.Contains(s.ProcessId)) list.Add(s.ProcessId);
        }

        int applied = 0;
        var skipped = new List<string>();

        foreach (var ov in profile.AppOverrides)
        {
            if (string.IsNullOrWhiteSpace(ov.ExePath)) continue;

            var appProfile = AppProfileService.Get(ov.AppProfileId);
            if (appProfile == null) continue;

            if (!runningByPath.TryGetValue(ov.ExePath, out var pids) || pids.Count == 0)
            {
                skipped.Add(Path.GetFileName(ov.ExePath));
                continue;
            }

            foreach (var pid in pids)
            {
                try
                {
                    AppAudioRoutingService.SetAppEndpoint(pid, DataFlow.Render, appProfile.OutputDeviceId);
                    AppAudioRoutingService.SetAppEndpoint(pid, DataFlow.Capture, appProfile.InputDeviceId);
                    applied++;
                }
                catch { }
            }
        }

        var voicemeeterStatus = profile.RestartVoicemeeterAfterApply
            ? VoicemeeterService.RestartAudioEngine()
            : VoicemeeterRestartStatus.NotRequested;

        return new ProfileApplyResult(applied, skipped, missing, voicemeeterStatus);
    }

    public static List<uint> GetRunningPidsForExe(string exePath)
    {
        var set = new HashSet<string>(new[] { exePath }, StringComparer.OrdinalIgnoreCase);
        var map = GetRunningProcessesByPath(set);
        return map.TryGetValue(exePath, out var list) ? list : new List<uint>();
    }

    // Only probes MainModule for processes whose name matches targets — skipping the rest
    // (MainModule access is slow; full-system scan takes several seconds).
    public static Dictionary<string, List<uint>> GetRunningProcessesByPath(HashSet<string> targetPaths)
    {
        var map = new Dictionary<string, List<uint>>(StringComparer.OrdinalIgnoreCase);
        if (targetPaths.Count == 0) return map;

        var targetNames = new HashSet<string>(
            targetPaths.Select(Path.GetFileNameWithoutExtension).Where(n => !string.IsNullOrEmpty(n))!,
            StringComparer.OrdinalIgnoreCase);

        foreach (var p in Process.GetProcesses())
        {
            try
            {
                if (!targetNames.Contains(p.ProcessName)) continue;
                var path = p.MainModule?.FileName;
                if (string.IsNullOrEmpty(path)) continue;
                if (!targetPaths.Contains(path)) continue;
                if (!map.TryGetValue(path, out var list)) map[path] = list = new List<uint>();
                list.Add((uint)p.Id);
            }
            catch { }
            finally { p.Dispose(); }
        }
        return map;
    }
}
