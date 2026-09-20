using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace AudioDeviceSwitcher;

public partial class ProfileScheduleEditDialog : Window
{
    private readonly ProfileScheduleService _service;
    private readonly Guid? _id;
    public Guid? SavedId { get; private set; }

    public ProfileScheduleEditDialog(ProfileScheduleService service, Guid? id = null)
    {
        InitializeComponent();
        _service = service;
        _id = id;
        string[] names = ["日", "一", "二", "三", "四", "五", "六"];
        foreach (int day in new[] { 1, 2, 3, 4, 5, 6, 0 })
            WeekdaysPanel.Children.Add(new CheckBox { Content = "周" + names[day], Tag = (DayOfWeek)day, Margin = new Thickness(0, 4, 8, 0) });
        ProfileChoice.ItemsSource = ProfileService.GetAll();
        OnceDate.SelectedDate = DateTime.Today;
        if (id is Guid editingId)
        {
            var item = service.Items.First(s => s.Id == editingId);
            Title = "编辑定时计划";
            SaveButton.Content = "保存修改";
            ProfileChoice.SelectedValue = item.ProfileId;
            TimeInput.Text = item.Time;
            EnabledCheck.IsChecked = item.Enabled;
            OnceDate.SelectedDate = item.OnceDate;
            RepeatChoice.SelectedIndex = item.Weekdays.Count == 0 ? 2 : item.Weekdays.Distinct().Count() == 7 ? 0 : 1;
            foreach (CheckBox box in WeekdaysPanel.Children) box.IsChecked = item.Weekdays.Contains((DayOfWeek)box.Tag);
        }
        Loaded += (_, _) => ProfileChoice.Focus();
    }

    private void Repeat_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (OnceDate == null || WeekdaysPanel == null) return;
        OnceDate.Visibility = RepeatChoice.SelectedIndex == 2 ? Visibility.Visible : Visibility.Collapsed;
        WeekdaysPanel.Visibility = RepeatChoice.SelectedIndex == 1 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (ProfileChoice.SelectedValue is not Guid profileId) throw new InvalidOperationException("请选择目标音频方案。");
            if (!TimeOnly.TryParseExact(TimeInput.Text.Trim(), "HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out var time))
                throw new InvalidOperationException("请输入 24 小时制时间，例如 09:00 或 18:30。");
            var days = RepeatChoice.SelectedIndex == 0 ? Enum.GetValues<DayOfWeek>().ToList() :
                RepeatChoice.SelectedIndex == 1 ? WeekdaysPanel.Children.Cast<CheckBox>().Where(b => b.IsChecked == true).Select(b => (DayOfWeek)b.Tag).ToList() : [];
            if (RepeatChoice.SelectedIndex == 1 && days.Count == 0) throw new InvalidOperationException("请至少选择一个星期。");
            if (RepeatChoice.SelectedIndex == 2 && (!OnceDate.SelectedDate.HasValue || OnceDate.SelectedDate.Value.Date.Add(time.ToTimeSpan()) <= DateTime.Now))
                throw new InvalidOperationException("单次计划请选择未来的日期和时间。");
            // Read the latest execution history; a schedule may fire while this dialog is open.
            var original = _service.Items.FirstOrDefault(s => s.Id == _id);
            if (_id.HasValue && original == null) throw new InvalidOperationException("此计划已删除，请重新新建。");
            var item = (original ?? new ProfileSchedule()) with
            {
                ProfileId = profileId, Time = time.ToString("HH:mm"), Weekdays = days,
                OnceDate = OnceDate.SelectedDate ?? DateTime.Today, Enabled = EnabledCheck.IsChecked == true
            };
            _service.Upsert(item);
            SavedId = item.Id;
            DialogResult = true;
        }
        catch (Exception ex) { Message.Text = ex.Message; }
    }
}
