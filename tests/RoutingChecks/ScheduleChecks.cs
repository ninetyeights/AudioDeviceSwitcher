using AudioDeviceSwitcher;
using System.IO;

internal static class ScheduleChecks
{
    public static void Render(string outputPath)
    {
        var app = new System.Windows.Application();
        app.Resources.MergedDictionaries.Add(new System.Windows.ResourceDictionary
        { Source = new Uri("pack://application:,,,/AudioDeviceSwitcher;component/Theme.xaml") });
        var folder = Path.Combine(Path.GetTempPath(), "CKit-SchedulePreview-" + Guid.NewGuid().ToString("N"));
        var path = Path.Combine(folder, "schedules.json");
        try
        {
            using var service = new ProfileScheduleService(_ => "测试", path, startTimer: false);
            service.Upsert(new ProfileSchedule { Weekdays = [DayOfWeek.Monday, DayOfWeek.Wednesday, DayOfWeek.Friday],
                LastResult = "已跳过：配置已锁定，请先解锁", LastRun = DateTime.Now });
            var window = new ProfileScheduleWindow(service);
            var content = (System.Windows.FrameworkElement)window.Content;
            content.Measure(new System.Windows.Size(820, 580));
            content.Arrange(new System.Windows.Rect(0, 0, 820, 580));
            content.UpdateLayout();
            var bitmap = new System.Windows.Media.Imaging.RenderTargetBitmap(820, 580, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
            var background = new System.Windows.Media.DrawingVisual();
            using (var drawing = background.RenderOpen())
                drawing.DrawRectangle(System.Windows.Media.Brushes.White, null, new System.Windows.Rect(0, 0, 820, 580));
            bitmap.Render(background);
            bitmap.Render(content);
            var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmap));
            using (var stream = File.Create(outputPath)) encoder.Save(stream);
            window.Close();
            var dialog = new ProfileScheduleEditDialog(service);
            ((System.Windows.Controls.ComboBox)dialog.FindName("RepeatChoice")).SelectedIndex = 1;
            content = (System.Windows.FrameworkElement)dialog.Content;
            content.Measure(new System.Windows.Size(490, 520));
            content.Arrange(new System.Windows.Rect(0, 0, 490, 520));
            content.UpdateLayout();
            bitmap = new System.Windows.Media.Imaging.RenderTargetBitmap(490, 520, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
            bitmap.Render(background);
            bitmap.Render(content);
            encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmap));
            using (var stream = File.Create(Path.ChangeExtension(outputPath, "editor.png"))) encoder.Save(stream);
            dialog.Close();
        }
        finally
        {
            foreach (var file in new[] { path, path + ".bak", path + ".tmp" })
                if (File.Exists(file)) File.Delete(file);
            if (Directory.Exists(folder)) Directory.Delete(folder);
            app.Shutdown();
        }
    }

    public static void Run()
    {
        CheckReleaseFeatures();
        CheckEffectiveTime();
        CheckRecovery();
        var monday = new DateTime(2026, 9, 21, 8, 59, 59);
        var daily = new ProfileSchedule { Weekdays = Enum.GetValues<DayOfWeek>().ToList() };
        Check(daily.NextAfter(monday) == monday.Date.AddHours(9), "daily next time");
        Check(daily.NextAfter(monday.Date.AddHours(9)) == monday.Date.AddDays(1).AddHours(9), "no catch-up at startup");
        var weekly = daily with { Weekdays = [DayOfWeek.Friday] };
        Check(weekly.NextAfter(monday) == monday.Date.AddDays(4).AddHours(9), "weekday selection");
        Check((daily with { Enabled = false }).NextAfter(monday) == null, "disabled schedule");
        Check((daily with { Time = "25:80" }).NextAfter(monday) == null, "invalid time");
        var once = new ProfileSchedule { OnceDate = monday.Date };
        Check(once.NextAfter(monday) == monday.Date.AddHours(9), "single occurrence");
        Check(once.NextAfter(monday.AddDays(1)) == null, "expired single occurrence");
        Check((daily with { LastOccurrence = monday.Date.AddHours(9) }).NextAfter(monday) == monday.Date.AddDays(1).AddHours(9), "clock rollback cannot repeat claimed occurrence");
        Check(!ProfileScheduleService.IsContinuous(monday, monday.AddMinutes(2)), "resume gap skips catch-up");
        Check(!ProfileScheduleService.IsContinuous(monday, monday.AddHours(-1)), "backward jump skipped");

        var folder = Path.Combine(Path.GetTempPath(), "CKit-ScheduleChecks-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "schedules.json");
        try
        {
            var now = monday;
            int calls = 0;
            using (var service = new ProfileScheduleService(_ => { calls++; return "已应用"; }, path, () => now, false))
            {
                service.Upsert(daily);
                service.Upsert(daily with { Id = Guid.NewGuid() });
                now = monday.AddSeconds(2);
                service.CheckNow();
                Check(calls == 1, "same-time conflict executes first only");
                Check(service.Items[1].LastResult.Contains("已跳过"), "conflict result saved");
                service.CheckNow();
                Check(calls == 1, "same tick cannot repeat");
                now = monday;
                service.CheckNow();
                now = monday.AddSeconds(2);
                service.CheckNow();
                Check(calls == 1, "clock rollback does not reapply");
                service.RunNow(daily.Id);
                Check(calls == 2 && service.Items[0].LastOccurrence == monday.Date.AddHours(9), "manual run preserves schedule claim");
                now = monday.AddDays(1).AddMinutes(-1);
                service.CheckNow();
                now = monday.AddDays(1).AddMinutes(1);
                service.CheckNow();
                Check(calls == 2, "missed time does not execute after gap");
            }
            now = monday;
            using var restarted = new ProfileScheduleService(_ => { calls++; return "已应用"; }, path, () => now, false);
            now = monday.AddSeconds(2);
            restarted.CheckNow();
            Check(calls == 2, "persisted claim prevents duplicate after restart");
            Check(restarted.Items.Count == 2, "plans persist");
            restarted.Delete(daily.Id);
            Check(restarted.Items.Count == 1, "delete plan");
        }
        finally
        {
            // Only remove fixture files created under this unique test directory.
            foreach (var file in new[] { path, path + ".bak", path + ".tmp" })
                if (File.Exists(file)) File.Delete(file);
            Directory.Delete(folder);
        }
    }

    private static void Check(bool passed, string name)
    {
        if (!passed) throw new Exception("FAIL " + name);
        Console.WriteLine("PASS " + name);
    }

    private static void CheckRecovery()
    {
        var folder = Path.Combine(Path.GetTempPath(), "CKit-RecoveryChecks-" + Guid.NewGuid().ToString("N"));
        var path = Path.Combine(folder, "schedules.json");
        var now = new DateTime(2026, 9, 21, 19, 0, 0);
        var work = new ProfileSchedule { ProfileId = Guid.NewGuid(), Time = "09:00", Weekdays = Enum.GetValues<DayOfWeek>().ToList(), EffectiveFrom = now.Date.AddDays(-1) };
        var play = work with { Id = Guid.NewGuid(), ProfileId = Guid.NewGuid(), Time = "18:00" };
        var calls = new List<Guid>();
        try
        {
            using (var service = new ProfileScheduleService(id => { calls.Add(id); return "完成"; }, path, () => now, false, () => true))
            {
                service.Upsert(work);
                service.Upsert(play);
                service.CheckNow();
                Check(calls.SequenceEqual(new[] { play.ProfileId }), "startup catches latest only");
                service.NotifyResumed();
                service.CheckNow();
                Check(calls.Count == 1, "handled latest never falls back to older missed plan");
                now = now.AddDays(1).Date.AddHours(10);
                service.NotifyResumed();
                service.CheckNow();
                Check(calls.Count == 2 && calls.Last() == work.ProfileId, "resume applies morning plan");
                service.Upsert(play with { Enabled = false });
                now = now.Date.AddHours(19);
                service.NotifyResumed();
                service.CheckNow();
                Check(calls.Count == 2, "disabled plan excluded from recovery");
            }
            using (var service = new ProfileScheduleService(id => { calls.Add(id); return "完成"; }, path, () => now, false, () => true))
            {
                service.CheckNow();
                Check(calls.Count == 2, "recovery claim survives restart");
                now = now.AddDays(1).Date.AddHours(9);
                var once = new ProfileSchedule { ProfileId = Guid.NewGuid(), OnceDate = now.Date, Time = "08:30", EffectiveFrom = now.Date };
                service.Upsert(work with { Enabled = false });
                service.Upsert(once);
                service.NotifyResumed();
                service.CheckNow();
                Check(calls.Count == 3 && calls.Last() == once.ProfileId, "missed single plan recoverable");
            }
            now = now.AddDays(1).Date.AddHours(8).AddMinutes(59).AddSeconds(55);
            using (var service = new ProfileScheduleService(id => { calls.Add(id); return "完成"; }, path, () => now, false))
            {
                service.Upsert(work);
                service.CheckNow();
                now = now.AddSeconds(10);
                service.NotifyResumed();
                service.CheckNow();
                Check(calls.Count == 3, "disabled recovery skips even short sleep");
            }
        }
        finally
        {
            foreach (var file in new[] { path, path + ".bak", path + ".tmp" })
                if (File.Exists(file)) File.Delete(file);
            if (Directory.Exists(folder)) Directory.Delete(folder);
        }
    }

    private static void CheckEffectiveTime()
    {
        var folder = Path.Combine(Path.GetTempPath(), "CKit-EffectiveChecks-" + Guid.NewGuid().ToString("N"));
        var path = Path.Combine(folder, "schedules.json");
        var now = new DateTime(2026, 9, 21, 10, 0, 0);
        int calls = 0;
        try
        {
            using (var service = new ProfileScheduleService(_ => { calls++; return "完成"; }, path, () => now, false, () => true))
            {
                var item = new ProfileSchedule { Time = "18:00", Weekdays = Enum.GetValues<DayOfWeek>().ToList() };
                service.Upsert(item);
                service.CheckNow();
                Check(calls == 0, "new evening plan cannot catch yesterday before first due time");
                service.NotifyResumed();
                service.CheckNow();
                Check(calls == 0, "waking before first due time does nothing");
                now = now.Date.AddHours(19);
                service.NotifyResumed();
                service.CheckNow();
                Check(calls == 1, "genuinely missed first occurrence is recovered");
                service.Upsert(service.Items[0] with { Time = "20:00" });
                service.NotifyResumed();
                service.CheckNow();
                Check(calls == 1, "editing to later time does not recover yesterday");
                service.Upsert(service.Items[0] with { Enabled = false });
                now = now.AddHours(2);
                service.Upsert(service.Items[0] with { Enabled = true });
                service.NotifyResumed();
                service.CheckNow();
                Check(calls == 1, "re-enable does not catch occurrences while disabled");
            }
            using (var service = new ProfileScheduleService(_ => { calls++; return "完成"; }, path, () => now, false, () => true))
            {
                service.CheckNow();
                Check(calls == 1, "effective time survives restart");
            }
            File.WriteAllText(path, System.Text.Json.JsonSerializer.Serialize(new[] {
                new ProfileSchedule { Time = "22:00", Weekdays = Enum.GetValues<DayOfWeek>().ToList() }
            }));
            using (var service = new ProfileScheduleService(_ => { calls++; return "完成"; }, path, () => now, false, () => true))
            {
                service.CheckNow();
                Check(calls == 1 && service.Items[0].EffectiveFrom == now, "legacy plans migrate without historical catch-up");
                now = now.Date.AddHours(21).AddMinutes(59).AddSeconds(59);
                service.CheckNow();
                now = now.AddSeconds(2);
                service.CheckNow();
                Check(calls == 2, "normal future occurrence still executes");
            }
        }
        finally
        {
            foreach (var file in new[] { path, path + ".bak", path + ".tmp" })
                if (File.Exists(file)) File.Delete(file);
            if (Directory.Exists(folder)) Directory.Delete(folder);
        }
    }

    private static void CheckReleaseFeatures()
    {
        var folder = Path.Combine(Path.GetTempPath(), "CKit-ReleaseChecks-" + Guid.NewGuid().ToString("N"));
        var path = Path.Combine(folder, "schedules.json");
        var now = new DateTime(2026, 9, 21, 8, 59, 59);
        var profileId = Guid.NewGuid();
        var item = new ProfileSchedule { ProfileId = profileId, Weekdays = Enum.GetValues<DayOfWeek>().ToList() };
        bool ready = false;
        try
        {
            using var service = new ProfileScheduleService(_ => "等待确认：已提交", path, () => now, false,
                verify: _ => ready ? ("成功", false) : ("等待会话", true));
            service.Upsert(item);
            now = now.AddSeconds(2);
            service.CheckNow();
            Check(service.Items[0].AwaitingVerification, "execution arms result verification");
            now = now.AddSeconds(6);
            service.CheckNow();
            Check(service.Items[0].LastResult == "等待会话", "pending execution displays actual waiting state");
            ready = true;
            now = now.AddSeconds(6);
            service.CheckNow();
            Check(service.Items[0].LastResult == "成功" && !service.Items[0].AwaitingVerification, "verification converges to success");
            var claimed = service.Items[0].LastOccurrence;
            service.RunNow(item.Id);
            Check(service.Items[0].AwaitingVerification && service.Items[0].LastOccurrence == claimed, "retry preserves scheduled occurrence");
            var backup = new BackupService.Backup { Profiles = [new DeviceProfile { Id = profileId }], Schedules = service.Items.ToList(), CatchUpProfileSchedules = true };
            var roundtrip = System.Text.Json.JsonSerializer.Deserialize<BackupService.Backup>(System.Text.Json.JsonSerializer.Serialize(backup))!;
            BackupService.ValidateSchedules(roundtrip, []);
            Check(roundtrip.Schedules.Count == 1 && roundtrip.CatchUpProfileSchedules == true, "backup roundtrip includes schedules and catch-up preference");
            service.MergeImported(roundtrip.Schedules);
            Check(service.Items.Count == 1 && service.Items[0].LastOccurrence == null && service.Items[0].EffectiveFrom == now
                && !service.Items[0].AwaitingVerification, "import merges by ID and resets execution history and effective time");
            Check(service.Items[0].LatestAt(now) == null, "import cannot catch old occurrences");
            var old = System.Text.Json.JsonSerializer.Deserialize<BackupService.Backup>("{\"FormatVersion\":1}")!;
            BackupService.ValidateSchedules(old, []);
            Check(old.Schedules.Count == 0 && old.CatchUpProfileSchedules == null, "v1 backup remains compatible");
            roundtrip.Profiles.Clear();
            bool rejected = false;
            try { BackupService.ValidateSchedules(roundtrip, []); } catch (InvalidDataException) { rejected = true; }
            Check(rejected, "missing schedule target rejected before import writes");
        }
        finally
        {
            foreach (var file in new[] { path, path + ".bak", path + ".tmp" })
                if (File.Exists(file)) File.Delete(file);
            if (Directory.Exists(folder)) Directory.Delete(folder);
        }
    }
}
