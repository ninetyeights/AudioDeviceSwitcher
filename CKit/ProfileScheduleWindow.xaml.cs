using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace AudioDeviceSwitcher;

public partial class ProfileScheduleWindow : Window
{
    private readonly ProfileScheduleService _service;
    private readonly DispatcherTimer _refresh = new() { Interval = TimeSpan.FromSeconds(15) };
    private Guid? _editingId;
    private bool _refreshing;
    private static readonly string[] DayNames = ["日", "一", "二", "三", "四", "五", "六"];
    private record Row(Guid Id, string State, string Profile, string Timing, string Next, string Result);

    public ProfileScheduleWindow(ProfileScheduleService service)
    {
        InitializeComponent();
        _service = service;
        CatchUpCheck.IsChecked = SettingsService.Load().CatchUpProfileSchedules;
        _service.Changed += Refresh;
        _refresh.Tick += (_, _) => Refresh();
        _refresh.Start();
        Closed += (_, _) => { _refresh.Stop(); _service.Changed -= Refresh; };
        Refresh();
        Message.Text = "点击「＋ 新建计划…」打开添加窗口；选中列表中的计划后可编辑。";
    }

    private void Refresh()
    {
        var profiles = ProfileService.GetAll();
        var now = DateTime.Now;
        _refreshing = true;
        var rows = _service.Items.Select(s => new Row(s.Id, s.Enabled ? "启用" : "停用",
            profiles.Find(p => p.Id == s.ProfileId)?.Name ?? "音频方案已删除",
            s.Time + " · " + (s.Weekdays.Count == 0 ? s.OnceDate.ToString("yyyy-MM-dd") :
                s.Weekdays.Distinct().Count() == 7 ? "每天" : string.Join("、", s.Weekdays.Select(d => "周" + DayNames[(int)d]))),
            !s.Enabled ? "—" : s.NextAfter(now)?.ToString("MM-dd HH:mm") ?? "已结束 / 已错过",
            (s.LastRun?.ToString("MM-dd HH:mm:ss") is string stamp ? stamp + "\n" : "") + s.LastResult)).ToList();
        ScheduleList.ItemsSource = rows;
        ScheduleList.SelectedItem = rows.Find(r => r.Id == _editingId);
        _refreshing = false;
        if (_service.Error != null) Message.Text = _service.Error;
    }

    private void CatchUp_Click(object sender, RoutedEventArgs e)
    {
        var settings = SettingsService.Load();
        bool previous = settings.CatchUpProfileSchedules;
        try
        {
            settings.CatchUpProfileSchedules = CatchUpCheck.IsChecked == true;
            SettingsService.Save();
            Message.Text = settings.CatchUpProfileSchedules
                ? "已开启：下次启动或唤醒时生效，本次不会立即切换。"
                : "已关闭：错过的计划不再补执行。";
        }
        catch (Exception ex)
        {
            settings.CatchUpProfileSchedules = previous;
            CatchUpCheck.IsChecked = previous;
            Message.Text = "保存失败：" + ex.Message;
        }
    }

    private void Selection_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_refreshing) return;
        _editingId = (ScheduleList.SelectedItem as Row)?.Id;
        UpdateButtons();
    }

    private void UpdateButtons()
    {
        EditButton.IsEnabled = ToggleButton.IsEnabled = RunButton.IsEnabled = DeleteButton.IsEnabled = _editingId.HasValue;
    }

    private void OpenEditor(Guid? id)
    {
        var dialog = new ProfileScheduleEditDialog(_service, id) { Owner = this };
        if (dialog.ShowDialog() != true) return;
        _editingId = dialog.SavedId;
        Refresh();
        UpdateButtons();
        Message.Text = id.HasValue ? "修改已保存。" : "计划已添加，请查看下次执行时间。";
    }

    private void New_Click(object sender, RoutedEventArgs e) => Guard(() => OpenEditor(null));
    private void Edit_Click(object sender, RoutedEventArgs e) => Guard(() =>
    {
        if (_editingId is Guid id) OpenEditor(id);
    });
    private void Toggle_Click(object sender, RoutedEventArgs e) => Guard(() =>
    {
        if (_editingId is not Guid id) return;
        var item = _service.Items.First(s => s.Id == id);
        _service.Upsert(item with { Enabled = !item.Enabled });

        Message.Text = item.Enabled ? "计划已停用。" : "计划已启用。";
    });

    private void Run_Click(object sender, RoutedEventArgs e) => Guard(() =>
    {
        if (_editingId is Guid id) { _service.RunNow(id); Message.Text = "执行已提交，请查看最近结果；原定时计划不受影响。"; }
    });

    private void Delete_Click(object sender, RoutedEventArgs e) => Guard(() =>
    {
        if (_editingId is Guid id) _service.Delete(id);
        _editingId = null;
        Refresh();
        UpdateButtons();
        Message.Text = "计划已删除。";
    });

    private void Guard(Action action)
    {
        try { action(); }
        catch (Exception ex) { Message.Text = ex.Message; }
    }
}
