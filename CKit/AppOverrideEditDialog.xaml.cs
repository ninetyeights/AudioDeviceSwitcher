using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace AudioDeviceSwitcher;

public partial class AppOverrideEditDialog : Window
{
    public record SessionOption(string ExePath, string DisplayName);

    public string ExePath { get; private set; } = "";
    public Guid AppProfileId { get; private set; }

    private readonly List<AppProfile> _profiles;

    public AppOverrideEditDialog(AppOverride? existing = null)
    {
        InitializeComponent();

        _profiles = AppProfileService.GetAll();
        AppProfileBox.ItemsSource = _profiles;

        if (_profiles.Count == 0)
        {
            AppProfileBox.IsEnabled = false;
        }

        // Session dropdown from currently-active audio apps that expose an exe path
        var sessions = AudioSessionService.GetActiveAppSessions()
            .Where(s => !string.IsNullOrEmpty(s.ExecutablePath))
            .Select(s => new SessionOption(s.ExecutablePath!, $"{s.DisplayName} ({Path.GetFileName(s.ExecutablePath)})"))
            .GroupBy(s => s.ExePath, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToList();
        SessionBox.ItemsSource = sessions;

        if (existing != null)
        {
            // Pre-fill for edit
            FromFileRadio.IsChecked = true;
            ExePathBox.Text = existing.ExePath;
            AppProfileBox.SelectedItem = _profiles.Find(p => p.Id == existing.AppProfileId);
        }
        else if (sessions.Count == 0)
        {
            FromFileRadio.IsChecked = true;
        }
    }

    private void SourceRadio_Changed(object sender, RoutedEventArgs e) { /* binding handles enablement */ }

    private void Browse_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "选择可执行文件",
            Filter = "可执行文件 (*.exe)|*.exe|所有文件 (*.*)|*.*",
            InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
        };
        if (dlg.ShowDialog() == true) ExePathBox.Text = dlg.FileName;
    }

    private void OK_Click(object sender, RoutedEventArgs e)
    {
        string path = "";
        if (FromSessionRadio.IsChecked == true)
        {
            if (SessionBox.SelectedItem is not SessionOption opt)
            {
                MessageBox.Show("请从下拉中选择一个应用。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            path = opt.ExePath;
        }
        else
        {
            path = ExePathBox.Text.Trim();
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                MessageBox.Show("请选择一个有效的可执行文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }

        if (AppProfileBox.SelectedItem is not AppProfile ap)
        {
            var msg = _profiles.Count == 0
                ? "还没有应用配置，请先在主窗口的\"应用配置\"里创建。"
                : "请选择一个应用配置。";
            MessageBox.Show(msg, "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        ExePath = path;
        AppProfileId = ap.Id;
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
}
