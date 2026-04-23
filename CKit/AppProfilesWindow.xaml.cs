using System.Windows;
using System.Windows.Controls;

namespace AudioDeviceSwitcher;

public partial class AppProfilesWindow : Window
{
    public AppProfilesWindow()
    {
        InitializeComponent();
        Loaded += (_, _) => LoadProfiles();
    }

    private void LoadProfiles()
    {
        var profiles = AppProfileService.GetAll();
        ProfileList.ItemsSource = profiles;
        EmptyHint.Visibility = profiles.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void New_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new AppProfileEditDialog { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            AppProfileService.Save(dlg.Result);
            LoadProfiles();
        }
    }

    private void Edit_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not Guid id) return;
        var profile = AppProfileService.Get(id);
        if (profile == null) return;

        var dlg = new AppProfileEditDialog(profile) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            AppProfileService.Save(dlg.Result);
            LoadProfiles();
        }
    }

    private void Delete_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not Guid id) return;
        var profile = AppProfileService.Get(id);
        if (profile == null) return;

        var confirm = MessageBox.Show(
            $"确定要删除应用配置「{profile.Name}」吗？\n使用该配置的系统配置中的相关应用覆盖将失效。",
            "确认删除", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        AppProfileService.Delete(id);
        LoadProfiles();
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
