using AudioDeviceSwitcher;
using NAudio.CoreAudioApi;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;
using System.IO;
using System.Text.Json;

internal static class LiveAcceptance
{
    public record Route(uint Pid, string? Path, string? Output, string? Input);
    public record Snapshot(string? Output, string? Input, List<Route> Routes);

    // Explicitly invoked manual acceptance only; ordinary regression runs never change routing.
    public static void Run(string action, string path)
    {
        if (action == "snapshot")
        {
            var routes = AudioSessionService.GetActiveAppSessions().Select(s => new Route(s.ProcessId, s.ExecutablePath,
                Read(s.ProcessId, DataFlow.Render), Read(s.ProcessId, DataFlow.Capture))).ToList();
            var snapshot = new Snapshot(AudioDeviceService.GetPlaybackDevices().Find(d => d.IsDefault)?.Id,
                AudioDeviceService.GetRecordingDevices().Find(d => d.IsDefault)?.Id, routes);
            File.WriteAllText(path, JsonSerializer.Serialize(snapshot, new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine($"Saved defaults and {routes.Count} session routes");
            return;
        }
        if (action == "restore-routes")
        {
            var snapshot = JsonSerializer.Deserialize<Snapshot>(File.ReadAllText(path))!;
            foreach (var route in snapshot.Routes)
            {
                // Refuse PID reuse; only restore the same still-running executable.
                if (route.Path == null || !ProfileApplyService.GetRunningPidsForExe(route.Path).Contains(route.Pid)) continue;
                AppAudioRoutingService.SetAppEndpoint(route.Pid, DataFlow.Render, route.Output);
                AppAudioRoutingService.SetAppEndpoint(route.Pid, DataFlow.Capture, route.Input);
                if (Read(route.Pid, DataFlow.Render) != route.Output || Read(route.Pid, DataFlow.Capture) != route.Input)
                    throw new Exception("Route restoration verification failed");
            }
            if (snapshot.Output != null) AudioDeviceService.SetDefaultDevice(snapshot.Output);
            if (snapshot.Input != null) AudioDeviceService.SetDefaultDevice(snapshot.Input);
            Console.WriteLine("PASS original device defaults and surviving app routes restored");
            return;
        }
        using var enumerator = new MMDeviceEnumerator();
        using var device = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
        using var output = new WasapiOut(device, AudioClientShareMode.Shared, true, 100);
        output.Init(new SignalGenerator { Gain = 0 }.ToWaveProvider());
        output.Play();
        uint pid = (uint)Environment.ProcessId;
        if (action == "seed")
        {
            var snapshot = JsonSerializer.Deserialize<Snapshot>(File.ReadAllText(path))!;
            AppAudioRoutingService.SetAppEndpoint(pid, DataFlow.Render, snapshot.Output);
            AppAudioRoutingService.SetAppEndpoint(pid, DataFlow.Capture, snapshot.Input);
            if (Read(pid, DataFlow.Render) != snapshot.Output || Read(pid, DataFlow.Capture) != snapshot.Input)
                throw new Exception("Seed did not persist");
            Console.WriteLine("PASS probe has explicit input/output routes; process now exits");
        }
        else if (action == "check-default")
        {
            if (Read(pid, DataFlow.Render) != null || Read(pid, DataFlow.Capture) != null)
                throw new Exception("Closed probe retained old route after profile switch");
            Console.WriteLine("PASS restarted probe uses follow-system for BOTH input and output");
            Console.WriteLine("Default output: " + device.FriendlyName);
        }
        else throw new ArgumentException("Unknown live acceptance action");
        output.Stop();
    }

    private static string? Read(uint pid, DataFlow flow)
    {
        var query = AppAudioRoutingService.QueryAppEndpoint(pid, flow);
        if (!query.Success) throw new Exception("Cannot read persisted route");
        return query.DeviceId;
    }
}
