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
    public static ProfileApplyResult Apply(DeviceProfile profile)
    {
        var allPlayback = AudioDeviceService.GetPlaybackDevices();
        var allRecording = AudioDeviceService.GetRecordingDevices();
        var currentPlayback = allPlayback.Find(d => d.IsDefault);
        var currentRecording = allRecording.Find(d => d.IsDefault);
        var currentPlaybackComm = AudioDeviceService.GetCommunicationsDefault(DataFlow.Render);
        var currentRecordingComm = AudioDeviceService.GetCommunicationsDefault(DataFlow.Capture);

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

        bool defaultsMatch = profile.PlaybackDeviceId == currentPlayback?.Id
                          && profile.RecordingDeviceId == currentRecording?.Id
                          && profile.PlaybackDeviceId == currentPlaybackComm.Id
                          && profile.RecordingDeviceId == currentRecordingComm.Id;
        // Skip only when system defaults match AND no AppOverride has drifted.
        // If anything drifted, re-run so user-triggered re-apply actually fixes them.
        if (defaultsMatch && !HasDriftedOverrides(profile))
            return new ProfileApplyResult(0, [], missing, VoicemeeterRestartStatus.NotRequested);

        if (playbackAvailable)
            AudioDeviceService.SetDefaultDevice(profile.PlaybackDeviceId!);
        if (recordingAvailable)
            AudioDeviceService.SetDefaultDevice(profile.RecordingDeviceId!);

        // Reset apps managed by other profiles but not by this one — back to follow-system.
        // Apps never referenced by any profile stay untouched.
        var thisProfileExes = profile.AppOverrides
            .Select(o => o.ExePath)
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var allManagedExes = ProfileService.GetAll()
            .SelectMany(p => p.AppOverrides)
            .Select(o => o.ExePath)
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var runningByPath = GetRunningProcessesByPath(allManagedExes);

        // Merge audio-session PIDs: Chrome-style apps often spawn child processes whose
        // session PID differs from the main-exe PID. Setting only on one leaves the other
        // unset — per-app override works (route changes) but queries via the other PID
        // return "follow system", confusing the UI.
        foreach (var s in AudioSessionService.GetActiveAppSessions())
        {
            if (string.IsNullOrEmpty(s.ExecutablePath)) continue;
            if (!allManagedExes.Contains(s.ExecutablePath)) continue;
            if (!runningByPath.TryGetValue(s.ExecutablePath, out var list))
                runningByPath[s.ExecutablePath] = list = new List<uint>();
            if (!list.Contains(s.ProcessId)) list.Add(s.ProcessId);
        }

        foreach (var exePath in allManagedExes)
        {
            if (thisProfileExes.Contains(exePath)) continue;
            if (!runningByPath.TryGetValue(exePath, out var pids)) continue;
            foreach (var pid in pids)
            {
                try
                {
                    AppAudioRoutingService.SetAppEndpoint(pid, DataFlow.Render, null);
                    AppAudioRoutingService.SetAppEndpoint(pid, DataFlow.Capture, null);
                }
                catch { }
            }
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

    public static bool HasDriftedOverrides(DeviceProfile profile)
    {
        var exes = profile.AppOverrides
            .Select(o => o.ExePath)
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToHashSet(StringComparer.OrdinalIgnoreCase)!;
        if (exes.Count == 0) return false;

        var runningByPath = GetRunningProcessesByPath(exes);

        foreach (var ov in profile.AppOverrides)
        {
            if (string.IsNullOrWhiteSpace(ov.ExePath)) continue;
            var appProfile = AppProfileService.Get(ov.AppProfileId);
            if (appProfile == null) continue;

            if (!runningByPath.TryGetValue(ov.ExePath, out var pids) || pids.Count == 0) continue;

            string? actualOut = null, actualIn = null;
            foreach (var pid in pids)
            {
                if (actualOut == null) { try { actualOut = AppAudioRoutingService.GetAppEndpoint(pid, DataFlow.Render); } catch { } }
                if (actualIn == null) { try { actualIn = AppAudioRoutingService.GetAppEndpoint(pid, DataFlow.Capture); } catch { } }
                if (actualOut != null && actualIn != null) break;
            }

            bool outOk = string.Equals(actualOut ?? "", appProfile.OutputDeviceId ?? "", StringComparison.OrdinalIgnoreCase);
            bool inOk = string.Equals(actualIn ?? "", appProfile.InputDeviceId ?? "", StringComparison.OrdinalIgnoreCase);
            if (!outOk || !inOk) return true;
        }
        return false;
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
