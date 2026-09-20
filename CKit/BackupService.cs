using System.IO;
using System.Text.Json;

namespace AudioDeviceSwitcher;

// Import/export of user configuration as a single JSON backup file.
// Bundles device profiles + app presets + device nicknames, because a profile's
// AppOverrides reference app presets by id — they must travel together to restore intact.
public static class BackupService
{
    public class Backup
    {
        public int FormatVersion { get; set; } = 2;
        public string? AppVersion { get; set; }
        public string? ExportedAt { get; set; }
        public List<DeviceProfile> Profiles { get; set; } = [];
        public List<AppProfile> AppProfiles { get; set; } = [];
        public Dictionary<string, string> DeviceNicknames { get; set; } = [];
        public List<ProfileSchedule> Schedules { get; set; } = [];
        public bool? CatchUpProfileSchedules { get; set; }
    }

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static void Export(string path, ProfileScheduleService schedules)
    {
        var backup = new Backup
        {
            AppVersion = typeof(BackupService).Assembly.GetName().Version?.ToString(),
            ExportedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
            Profiles = ProfileService.GetAll(),
            AppProfiles = AppProfileService.GetAll(),
            DeviceNicknames = new Dictionary<string, string>(SettingsService.Load().DeviceNicknames),
            Schedules = schedules.Items.ToList(),
            CatchUpProfileSchedules = SettingsService.Load().CatchUpProfileSchedules,
        };
        File.WriteAllText(path, JsonSerializer.Serialize(backup, JsonOptions));
    }

    public static Backup ReadForImport(string path)
    {
        var json = File.ReadAllText(path);
        using var document = JsonDocument.Parse(json);
        if (document.RootElement.ValueKind != JsonValueKind.Object ||
            !document.RootElement.TryGetProperty(nameof(Backup.Profiles), out _) ||
            !document.RootElement.TryGetProperty(nameof(Backup.AppProfiles), out _))
            throw new InvalidDataException("这不是有效的音频方案备份文件");
        var backup = JsonSerializer.Deserialize<Backup>(json, JsonOptions)
            ?? throw new InvalidDataException("文件内容为空或格式不正确");
        ValidateForImport(backup);
        return backup;
    }

    private static void ValidateForImport(Backup backup)
    {
        if (backup.FormatVersion is < 1 or > 2)
            throw new InvalidDataException("无法识别的备份文件版本");
        ValidateSchedules(backup, ProfileService.GetAll().Select(p => p.Id));
        if (backup.Profiles.Any(p => p.Id == Guid.Empty || string.IsNullOrWhiteSpace(p.Name)) ||
            backup.AppProfiles.Any(p => p == null || p.Id == Guid.Empty || string.IsNullOrWhiteSpace(p.Name)))
            throw new InvalidDataException("音频方案或音频预设缺少有效的标识或名称");
        if (backup.Profiles.Select(p => p.Id).Distinct().Count() != backup.Profiles.Count ||
            backup.AppProfiles.Select(p => p.Id).Distinct().Count() != backup.AppProfiles.Count)
            throw new InvalidDataException("备份包含重复的音频方案或音频预设");
        var presetIds = AppProfileService.GetAll().Select(p => p.Id)
            .Concat(backup.AppProfiles.Select(p => p.Id)).ToHashSet();
        foreach (var profile in backup.Profiles)
        {
            if (profile.AppOverrides == null || profile.AppOverrides.Any(o => o == null ||
                string.IsNullOrWhiteSpace(o.ExePath) || !presetIds.Contains(o.AppProfileId)))
                throw new InvalidDataException($"音频方案“{profile.Name}”的应用设备规则无效或引用的音频预设不存在");
        }
    }

    public static string DescribeImport(Backup backup, ProfileScheduleService schedules)
    {
        string Count(string label, IEnumerable<Guid> incoming, IEnumerable<Guid> current)
        {
            var ids = incoming.ToList();
            var existing = current.ToHashSet();
            int replaced = ids.Count(existing.Contains);
            return $"{label}：新增 {ids.Count - replaced} 个，覆盖 {replaced} 个";
        }
        var nicknames = SettingsService.Load().DeviceNicknames;
        int replacedNames = backup.DeviceNicknames.Keys.Count(nicknames.ContainsKey);
        return "此备份用于当前电脑恢复。即将合并以下内容：\n\n" +
            Count("音频方案", backup.Profiles.Select(p => p.Id), ProfileService.GetAll().Select(p => p.Id)) + "\n" +
            Count("音频预设", backup.AppProfiles.Select(p => p.Id), AppProfileService.GetAll().Select(p => p.Id)) + "\n" +
            $"设备别名：新增 {backup.DeviceNicknames.Count - replacedNames} 个，覆盖 {replacedNames} 个\n" +
            Count("定时计划", backup.Schedules.Select(s => s.Id), schedules.Items.Select(s => s.Id)) +
            (backup.CatchUpProfileSchedules.HasValue ? $"\n错过计划补执行：{(backup.CatchUpProfileSchedules.Value ? "开启" : "关闭")}" : "") +
            "\n\n相同标识的项目会被覆盖，其余现有项目保留。\n未连接设备的设置原样保留。\n导入不会立即切换音频；计划保留启停状态，仅处理导入后的时间点。\n\n是否继续导入？";
    }

    public static (int Profiles, int AppProfiles, int Nicknames, int Schedules) ImportMerge(string path, ProfileScheduleService schedules)
        => ImportMerge(ReadForImport(path), schedules);

    // Apply the exact data shown in the preview, without rereading the source file.
    public static (int Profiles, int AppProfiles, int Nicknames, int Schedules) ImportMerge(Backup backup, ProfileScheduleService schedules)
    {
        ValidateForImport(backup);

        // App presets first, so profiles referencing them resolve after import.
        foreach (var ap in backup.AppProfiles)
            AppProfileService.Save(ap);

        var profiles = ProfileService.GetAll();
        foreach (var p in backup.Profiles)
        {
            var i = profiles.FindIndex(x => x.Id == p.Id);
            if (i >= 0) profiles[i] = p;
            else profiles.Add(p);
        }
        ProfileService.SaveAll(profiles);

        if (backup.DeviceNicknames.Count > 0)
        {
            var settings = SettingsService.Load();
            foreach (var kv in backup.DeviceNicknames)
                settings.DeviceNicknames[kv.Key] = kv.Value;
            SettingsService.Save();
        }

        schedules.MergeImported(backup.Schedules);
        if (backup.CatchUpProfileSchedules.HasValue)
        {
            SettingsService.Load().CatchUpProfileSchedules = backup.CatchUpProfileSchedules.Value;
            SettingsService.Save();
        }
        return (backup.Profiles.Count, backup.AppProfiles.Count, backup.DeviceNicknames.Count, backup.Schedules.Count);
    }

    public static void ValidateSchedules(Backup backup, IEnumerable<Guid> existingProfileIds)
    {
        if (backup.Profiles == null || backup.AppProfiles == null || backup.DeviceNicknames == null || backup.Schedules == null)
            throw new InvalidDataException("备份缺少有效的数据列表");
        if (backup.Profiles.Any(p => p == null) || backup.Schedules.Any(s => s == null || s.Id == Guid.Empty))
            throw new InvalidDataException("备份包含无效的音频方案或定时计划");
        var ids = existingProfileIds.Concat(backup.Profiles.Select(p => p.Id)).ToHashSet();
        if (backup.Schedules.Select(s => s.Id).Distinct().Count() != backup.Schedules.Count)
            throw new InvalidDataException("备份包含重复的定时计划");
        foreach (var item in backup.Schedules)
        {
            if (!ids.Contains(item.ProfileId)) throw new InvalidDataException("定时计划引用的音频方案不存在");
            if (item.Weekdays == null || item.Weekdays.Any(d => !Enum.IsDefined(d)) ||
                !TimeOnly.TryParseExact(item.Time, "HH:mm", System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out _))
                throw new InvalidDataException("定时计划的时间或重复日期无效");
        }
    }
}
