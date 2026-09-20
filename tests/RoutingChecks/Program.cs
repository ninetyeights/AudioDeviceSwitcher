using AudioDeviceSwitcher;
using NAudio.CoreAudioApi;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;
using System.Windows.Threading;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        if (args.Length == 3 && args[0] == "--live") { LiveAcceptance.Run(args[1], args[2]); return; }
        if (args.Length == 3 && args[0] == "--ui") { UiChecks.Render(args[1], args[2]); return; }
        if (args.Length == 2 && args[0] == "--schedule-ui") { ScheduleChecks.Render(args[1]); return; }
        ScheduleChecks.Run();
        var cases = new (string Name, bool Session, bool Success, string? Actual, string? Expected, string? Default, AppRouteState State)[]
        {
            ("No session is waiting, not drift", false, true, null, "headphones", "speaker", AppRouteState.WaitingForSession),
            ("Query failure is unknown", true, false, null, "headphones", "speaker", AppRouteState.Unknown),
            ("Query failure is not follow-system", true, false, null, null, "speaker", AppRouteState.Unknown),
            ("Both follow system", true, true, null, null, "speaker", AppRouteState.Matched),
            ("Follow-system resolves to expected device", true, true, null, "speaker", "speaker", AppRouteState.Matched),
            ("Expected follow-system resolves to actual device", true, true, "speaker", null, "speaker", AppRouteState.Matched),
            ("Different concrete devices drift", true, true, "speaker", "headphones", "speaker", AppRouteState.Drifted),
            ("Follow-system can genuinely drift", true, true, null, "headphones", "speaker", AppRouteState.Drifted),
            ("Unavailable default is unknown", true, true, null, "headphones", null, AppRouteState.Unknown),
            ("Device comparison ignores case", true, true, "DEVICE", "device", null, AppRouteState.Matched),
        };
        foreach (var c in cases)
        {
            var actual = AppRouteStatus.Evaluate(c.Session, c.Success, c.Actual, c.Expected, c.Default);
            if (actual != c.State) throw new Exception($"{c.Name}: expected {c.State}, got {actual}");
            Console.WriteLine($"PASS {c.Name}");
        }

        if (!args.Contains("--audio")) return;

        var dispatcher = Dispatcher.CurrentDispatcher;
        dispatcher.BeginInvoke(async () =>
        {
            try { await CheckNotifications(dispatcher); }
            catch (Exception ex) { Console.Error.WriteLine(ex); Environment.ExitCode = 1; }
            finally { dispatcher.BeginInvokeShutdown(DispatcherPriority.Background); }
        });
        Dispatcher.Run();
    }

    private static async Task CheckNotifications(Dispatcher dispatcher)
    {
        using var enumerator = new MMDeviceEnumerator();
        using var device = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
        var ready = new TaskCompletionSource();
        var created = new TaskCompletionSource();
        AudioSessionNotifier? notifier = null;
        using (notifier = new AudioSessionNotifier(dispatcher, () =>
        {
            ready.TrySetResult();
            if (notifier!.Snapshot.Any(s => s.ProcessId == (uint)Environment.ProcessId && s.Flow == DataFlow.Render))
                created.TrySetResult();
        }))
        {
            await ready.Task.WaitAsync(TimeSpan.FromSeconds(10));
            // Silence creates a real WASAPI session without changing the user's routing or volume.
            using var output = new WasapiOut(device, AudioClientShareMode.Shared, true, 100);
            output.Init(new SignalGenerator { Gain = 0 }.ToWaveProvider());
            output.Play();
            await created.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Console.WriteLine("PASS real WASAPI session creation notification (no polling)");
            output.Stop();
        }
        Console.WriteLine("PASS notification subscriptions disposed");
    }
}
