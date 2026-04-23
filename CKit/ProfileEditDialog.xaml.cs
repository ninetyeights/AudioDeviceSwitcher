using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace AudioDeviceSwitcher;

public partial class ProfileEditDialog : Window
{
    public record OverrideRow(string ExePath, string ExeDisplay, string ProfileDisplay);

    public string ProfileName => NameBox.Text.Trim();
    public int HotkeyModifiers { get; private set; }
    public int HotkeyKey { get; private set; }
    public List<AppOverride> AppOverrides { get; private set; } = [];

    public ProfileEditDialog(
        string? existingName = null,
        int modifiers = 0,
        int key = 0,
        List<AppOverride>? existingOverrides = null)
    {
        InitializeComponent();
        if (existingName != null)
            NameBox.Text = existingName;

        HotkeyModifiers = modifiers;
        HotkeyKey = key;
        UpdateHotkeyDisplay();

        if (existingOverrides != null)
            AppOverrides = existingOverrides.Select(o => new AppOverride { ExePath = o.ExePath, AppProfileId = o.AppProfileId }).ToList();

        RefreshOverrideList();

        Loaded += (_, _) => { NameBox.Focus(); NameBox.SelectAll(); };
    }

    private void RefreshOverrideList()
    {
        var profiles = AppProfileService.GetAll();
        var rows = AppOverrides.Select(o =>
        {
            var p = profiles.Find(x => x.Id == o.AppProfileId);
            var exeDisplay = string.IsNullOrEmpty(o.ExePath) ? "(未设置)" : Path.GetFileName(o.ExePath);
            var profileDisplay = p == null
                ? "(应用配置已删除)"
                : $"→ {p.Name}";
            return new OverrideRow(o.ExePath, exeDisplay, profileDisplay);
        }).ToList();
        OverrideList.ItemsSource = rows;
        EmptyOverrideHint.Visibility = rows.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void AddOverride_Click(object sender, RoutedEventArgs e)
    {
        if (AppProfileService.GetAll().Count == 0)
        {
            MessageBox.Show("还没有应用配置。请先在主窗口的\"应用配置\"里创建至少一个应用配置。",
                "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new AppOverrideEditDialog { Owner = this };
        if (dlg.ShowDialog() != true) return;

        // Dedup: replace existing entry for same path
        AppOverrides.RemoveAll(o => string.Equals(o.ExePath, dlg.ExePath, StringComparison.OrdinalIgnoreCase));
        AppOverrides.Add(new AppOverride { ExePath = dlg.ExePath, AppProfileId = dlg.AppProfileId });
        RefreshOverrideList();
    }

    private void EditOverride_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not string exePath) return;
        var existing = AppOverrides.Find(o => string.Equals(o.ExePath, exePath, StringComparison.OrdinalIgnoreCase));
        if (existing == null) return;

        var dlg = new AppOverrideEditDialog(existing) { Owner = this };
        if (dlg.ShowDialog() != true) return;

        AppOverrides.RemoveAll(o => string.Equals(o.ExePath, existing.ExePath, StringComparison.OrdinalIgnoreCase)
                                 || string.Equals(o.ExePath, dlg.ExePath, StringComparison.OrdinalIgnoreCase));
        AppOverrides.Add(new AppOverride { ExePath = dlg.ExePath, AppProfileId = dlg.AppProfileId });
        RefreshOverrideList();
    }

    private void DeleteOverride_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not string exePath) return;
        AppOverrides.RemoveAll(o => string.Equals(o.ExePath, exePath, StringComparison.OrdinalIgnoreCase));
        RefreshOverrideList();
    }

    private void HotkeyBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        e.Handled = true;

        var key = e.Key == Key.System ? e.SystemKey : e.Key;

        if (key is Key.LeftCtrl or Key.RightCtrl or Key.LeftAlt or Key.RightAlt
            or Key.LeftShift or Key.RightShift or Key.LWin or Key.RWin)
            return;

        int mod = 0;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Alt)) mod |= 0x0001;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control)) mod |= 0x0002;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)) mod |= 0x0004;

        if (mod == 0) return;

        HotkeyModifiers = mod;
        HotkeyKey = KeyInterop.VirtualKeyFromKey(key);
        UpdateHotkeyDisplay();
    }

    private void UpdateHotkeyDisplay()
    {
        if (HotkeyKey == 0)
        {
            HotkeyBox.Text = "点击设置快捷键";
            HotkeyBox.Foreground = Brushes.Gray;
            return;
        }

        var parts = new List<string>();
        if ((HotkeyModifiers & 0x0002) != 0) parts.Add("Ctrl");
        if ((HotkeyModifiers & 0x0001) != 0) parts.Add("Alt");
        if ((HotkeyModifiers & 0x0004) != 0) parts.Add("Shift");
        parts.Add(KeyInterop.KeyFromVirtualKey(HotkeyKey).ToString());

        HotkeyBox.Text = string.Join(" + ", parts);
        HotkeyBox.Foreground = Brushes.Black;
    }

    private void HotkeyBox_GotFocus(object sender, RoutedEventArgs e)
    {
        if (HotkeyKey == 0)
        {
            HotkeyBox.Text = "请按下快捷键组合...";
            HotkeyBox.Foreground = Brushes.Gray;
        }
    }

    private void HotkeyBox_LostFocus(object sender, RoutedEventArgs e) => UpdateHotkeyDisplay();

    private void ClearHotkey_Click(object sender, RoutedEventArgs e)
    {
        HotkeyModifiers = 0;
        HotkeyKey = 0;
        UpdateHotkeyDisplay();
    }

    private void OK_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(NameBox.Text))
        {
            MessageBox.Show("请输入配置名称。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
}
