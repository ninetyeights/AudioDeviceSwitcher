using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Windows.Threading;
using Microsoft.Win32;

namespace AudioDeviceSwitcher;

public record ProfileSchedule
{
    public Guid Id { get; init; } = Guid.NewGuid();
    public Guid ProfileId { get; set; }
    public bool Enabled { get; set; } = true;
    public string Time { get; set; } = "09:00";
    // Empty weekdays means a single occurrence on OnceDate.
    public List<DayOfWeek> Weekdays { get; set; } = [];
    public DateTime OnceDate { get; set; } = DateTime.Today;
    public DateTime? LastOccurrence { get; set; }
    public DateTime? LastRun { get; set; }
    public string LastResult { get; set; } = "尚未执行";
    // Only occurrences after creation, a timing edit, or re-enabling are eligible.
    public DateTime? EffectiveFrom { get; set; }
    public bool AwaitingVerification { get; set; }

    public DateTime? LatestAt(DateTime now)
    {
        if (!Enabled || !TimeOnly.TryParseExact(Time, "HH:mm", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var time)) return null;
        for (int offset = 0; offset <= 7; offset++)
        {
            var date = Weekdays.Count == 0 ? OnceDate.Date : now.Date.AddDays(-offset);
            var candidate = date.Add(time.ToTimeSpan());
            if (candidate <= now && (!EffectiveFrom.HasValue || candidate > EffectiveFrom.Value)
                && (Weekdays.Count == 0 || Weekdays.Contains(date.DayOfWeek))
                && !TimeZoneInfo.Local.IsInvalidTime(candidate)) return candidate;
            if (Weekdays.Count == 0) break;
        }
        return null;
    }

    public DateTime? NextAfter(DateTime after)
    {
        if (!Enabled || !TimeOnly.TryParseExact(Time, "HH:mm", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var time)) return null;
        for (int offset = 0; offset <= 7; offset++)
        {
            var date = Weekdays.Count == 0 ? OnceDate.Date : after.Date.AddDays(offset);
            var candidate = date.Add(time.ToTimeSpan());
            if (candidate > after && (!EffectiveFrom.HasValue || candidate > EffectiveFrom.Value)
                && (!LastOccurrence.HasValue || candidate > LastOccurrence.Value)
                && (Weekdays.Count == 0 || Weekdays.Contains(date.DayOfWeek))
                && !TimeZoneInfo.Local.IsInvalidTime(candidate)) return candidate;
            if (Weekdays.Count == 0) break;
        }
        return null;
    }
}

public sealed class ProfileScheduleService : IDisposable
{
    private static readonly string DefaultStorePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AudioDeviceSwitcher", "schedules.json");
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private List<ProfileSchedule> _items;
    private readonly DispatcherTimer _timer = new() { Interval = TimeSpan.FromSeconds(1) };
    private readonly Func<Guid, string> _apply;
    private readonly string _storePath;
    private readonly Func<DateTime> _clock;
    private DateTime _previous;
    private volatile bool _powerChanged;
    private volatile bool _suspended;
    private bool _startup = true;
    private readonly Func<bool> _catchUpEnabled;
    private readonly bool _started;
    private readonly Func<Guid, (string Message, bool Pending)>? _verify;
    private DateTime _lastVerification;
    public event Action? Changed;
    public IReadOnlyList<ProfileSchedule> Items => _items;
    public string? Error { get; private set; }

    public ProfileScheduleService(Func<Guid, string> apply, string? storePath = null,
        Func<DateTime>? clock = null, bool startTimer = true, Func<bool>? catchUpEnabled = null,
        Func<Guid, (string Message, bool Pending)>? verify = null)
    {
        _apply = apply;
        _verify = verify;
        _storePath = storePath ?? DefaultStorePath;
        _clock = clock ?? (() => DateTime.Now);
        _catchUpEnabled = catchUpEnabled ?? (() => false);
        _previous = _clock();
        _items = JsonStore.Read<List<ProfileSchedule>>(_storePath, JsonOptions) ?? [];
        foreach (var item in _items)
        {
            item.Weekdays ??= [];
            if (item.AwaitingVerification) item.LastResult = "检查已中断：程序曾退出，可立即执行 / 重试";
            item.AwaitingVerification = false;
        }
        // Older files have no reliable creation time. Start them from this upgrade,
        // rather than inventing historical occurrences that could switch devices now.
        if (_items.Any(s => !s.EffectiveFrom.HasValue))
        {
            foreach (var item in _items.Where(s => !s.EffectiveFrom.HasValue)) item.EffectiveFrom = _previous;
            try { Save(_items); }
            catch (Exception ex) { Error = "计划生效时间保存失败：" + ex.Message; }
        }
        _timer.Tick += (_, _) => CheckNow();
        _started = startTimer;
        if (startTimer)
        {
            SystemEvents.PowerModeChanged += OnPowerChanged;
            _timer.Start();
        }
    }

    private void Save(List<ProfileSchedule> items)
    {
        JsonStore.WriteAtomic(_storePath, JsonSerializer.Serialize(items, JsonOptions));
        _items = items;
        Error = null;
        Changed?.Invoke();
    }

    public void Upsert(ProfileSchedule item)
    {
        var items = _items.ToList();
        var index = items.FindIndex(s => s.Id == item.Id);
        var original = index >= 0 ? items[index] : null;
        bool changed = original != null && (original.Time != item.Time
            || original.ProfileId != item.ProfileId || original.OnceDate.Date != item.OnceDate.Date
            || !original.Weekdays.ToHashSet().SetEquals(item.Weekdays)
            || (!original.Enabled && item.Enabled));
        if (changed)
            item = item with { EffectiveFrom = _clock(), LastOccurrence = null, LastRun = null, AwaitingVerification = false, LastResult = "计划已更新，等待下一次执行" };
        else if (!item.EffectiveFrom.HasValue)
            item = item with { EffectiveFrom = _clock() };
        if (index < 0) items.Add(item); else items[index] = item;
        Save(items);
    }

    public void Delete(Guid id) => Save(_items.Where(s => s.Id != id).ToList());

    public void MergeImported(IEnumerable<ProfileSchedule> imported)
    {
        var items = _items.ToList();
        foreach (var schedule in imported)
        {
            var restored = schedule with { Weekdays = schedule.Weekdays.ToList(), EffectiveFrom = _clock(),
                LastOccurrence = null, LastRun = null, AwaitingVerification = false, LastResult = "已导入，等待下一次执行" };
            var index = items.FindIndex(s => s.Id == restored.Id);
            if (index < 0) items.Add(restored); else items[index] = restored;
        }
        Save(items);
    }

    public void RunNow(Guid id)
    {
        var item = _items.First(s => s.Id == id);
        // Manual testing does not consume the scheduled occurrence.
        var result = Execute(item.ProfileId);
        Upsert(item with { LastRun = _clock(), LastResult = "手动试运行：" + result,
            AwaitingVerification = NeedsVerification(result) });
    }

    private bool NeedsVerification(string result) => _verify != null && result.StartsWith("等待确认：", StringComparison.Ordinal);

    private void VerifyPending(DateTime now)
    {
        if (_verify == null || now - _lastVerification < TimeSpan.FromSeconds(5)) return;
        _lastVerification = now;
        foreach (var item in _items.Where(s => s.AwaitingVerification).ToArray())
        {
            try
            {
                var result = _verify(item.ProfileId);
                if (result.Message != item.LastResult || result.Pending != item.AwaitingVerification)
                    Upsert(item with { LastResult = result.Message, AwaitingVerification = result.Pending });
            }
            catch (Exception ex) { Error = "执行结果确认失败：" + ex.Message; Changed?.Invoke(); }
        }
    }

    private string Execute(Guid profileId)
    {
        try { return _apply(profileId); }
        catch (Exception ex) { return "执行失败：" + ex.Message; }
    }

    public static bool IsContinuous(DateTime previous, DateTime now) =>
        now >= previous && now - previous <= TimeSpan.FromSeconds(30);

    public void CheckNow()
    {
        var now = _clock();
        var previous = _previous;
        _previous = now;
        if (_suspended) return;
        VerifyPending(now);
        bool resumed = _powerChanged;
        bool recover = _startup || resumed;
        _startup = false;
        _powerChanged = false;
        if (recover && _catchUpEnabled()) { CatchUp(now); return; }
        if (resumed) return;
        // With recovery disabled, preserve the original no-catch-up behavior.
        if (!IsContinuous(previous, now)) return;
        var due = _items.Select(s => (Schedule: s, At: s.NextAfter(previous)))
            .Where(s => s.At.HasValue && s.At.Value <= now).ToList();
        var claimed = new HashSet<DateTime>();
        try
        {
            foreach (var (item, at) in due)
            {
                bool conflict = !claimed.Add(at!.Value);
                var recorded = item with
                {
                    LastOccurrence = at,
                    LastRun = now,
                    LastResult = conflict ? "已跳过：同一时间优先执行列表中靠前的计划" : "已开始执行，结果尚未确认"
                };
                // Persist the claim before touching devices; a restart cannot execute it twice.
                Upsert(recorded);
                if (!conflict)
                {
                    var result = Execute(item.ProfileId);
                    Upsert(recorded with { LastResult = result, AwaitingVerification = NeedsVerification(result) });
                }
            }
        }
        catch (Exception ex)
        {
            Error = "计划记录保存失败，本次执行已停止：" + ex.Message;
            Changed?.Invoke();
        }
    }

    private void OnPowerChanged(object sender, PowerModeChangedEventArgs e)
    {
        if (e.Mode == PowerModes.Suspend) _suspended = true;
        if (e.Mode == PowerModes.Resume) NotifyResumed();
    }

    public void NotifyResumed()
    {
        _suspended = false;
        _powerChanged = true;
    }

    private void CatchUp(DateTime now)
    {
        // Select before checking the claim: an already handled newer plan must
        // never allow an older missed plan to override it. OrderBy is stable for ties.
        var latest = _items.Select(s => (Schedule: s, At: s.LatestAt(now)))
            .Where(s => s.At.HasValue).OrderByDescending(s => s.At).FirstOrDefault();
        if (latest.Schedule == null || latest.Schedule.LastOccurrence >= latest.At) return;
        try
        {
            var recorded = latest.Schedule with
            {
                LastOccurrence = latest.At, LastRun = now,
                LastResult = "补执行已开始，结果尚未确认"
            };
            Upsert(recorded);
            var result = Execute(recorded.ProfileId);
            Upsert(recorded with { LastResult = $"补执行（原定 {latest.At:MM-dd HH:mm}）：" + result,
                AwaitingVerification = NeedsVerification(result) });
        }
        catch (Exception ex)
        {
            Error = "补执行记录保存失败，本次执行已停止：" + ex.Message;
            Changed?.Invoke();
        }
    }

    public void Dispose()
    {
        _timer.Stop();
        if (_started) SystemEvents.PowerModeChanged -= OnPowerChanged;
    }
}
