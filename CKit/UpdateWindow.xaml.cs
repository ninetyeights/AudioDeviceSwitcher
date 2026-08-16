using System.Diagnostics;
using System.Windows;

namespace AudioDeviceSwitcher;

public partial class UpdateWindow : Window
{
    private readonly UpdateInfo _info;
    private bool _busy;

    public UpdateWindow(UpdateInfo info)
    {
        InitializeComponent();
        _info = info;
        VersionText.Text = $"当前版本 {UpdateService.GetCurrentVersion().ToString(3)} → 最新版本 {info.Version.ToString(3)}";
        NotesText.Text = string.IsNullOrWhiteSpace(info.ReleaseNotes) ? "（无更新说明）" : info.ReleaseNotes;
    }

    private void Skip_Click(object sender, RoutedEventArgs e)
    {
        var s = SettingsService.Load();
        s.SkippedUpdateVersion = _info.Version.ToString(3);
        SettingsService.Save();
        Close();
    }

    private void Later_Click(object sender, RoutedEventArgs e) => Close();

    private void OpenBrowser_Click(object sender, RoutedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo(_info.HtmlUrl) { UseShellExecute = true }); }
        catch (Exception ex)
        {
            MessageBox.Show($"打开浏览器失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void Update_Click(object sender, RoutedEventArgs e)
    {
        if (_busy) return;
        if (string.IsNullOrEmpty(_info.InstallerAssetUrl))
        {
            MessageBox.Show("此版本未提供安装包，请点击「在浏览器中查看」手动下载。", "提示",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SetBusy(true);
        try
        {
            var progress = new Progress<double>(p =>
            {
                DownloadProgress.Value = p;
                ProgressText.Text = $"下载中… {p:P0}";
            });

            var path = await UpdateService.DownloadInstallerAsync(_info, progress, default);

            ProgressText.Text = "下载完成，正在启动安装程序…";
            // Requires admin (installer PrivilegesRequired=admin) — Process.Start blocks on
            // the UAC prompt and throws if the user declines it, which the catch below handles.
            UpdateService.LaunchInstaller(path);

            // The installer (CloseApplications=yes) would close this app itself once it
            // starts, but exiting explicitly avoids a Restart Manager prompt naming it.
            (Application.Current as App)?.ExitFromUI();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"下载或启动安装程序失败：\n{ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
            SetBusy(false);
        }
    }

    private void SetBusy(bool busy)
    {
        _busy = busy;
        UpdateButton.IsEnabled = !busy;
        SkipButton.IsEnabled = !busy;
        LaterButton.IsEnabled = !busy;
        ProgressPanel.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
    }
}
