using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;

namespace AudioDeviceSwitcher;

public partial class AboutWindow : Window
{
    public AboutWindow()
    {
        InitializeComponent();

        var asm = Assembly.GetExecutingAssembly();
        var ver = asm.GetName().Version?.ToString(3) ?? "1.0.0";
        VersionText.Text = $"版本 {ver}";

        var copyrightAttr = asm.GetCustomAttribute<AssemblyCopyrightAttribute>();
        CopyrightText.Text = copyrightAttr?.Copyright ?? "Copyright © Chester";
    }

    private void OpenLog_Click(object sender, RoutedEventArgs e)
    {
        var path = CrashLogger.LogFilePath;
        try
        {
            if (!File.Exists(path))
            {
                MessageBox.Show("\u6682\u65E0\u5D29\u6E83\u65E5\u5FD7\u3002",
                    "\u63D0\u793A", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"\u6253\u5F00\u65E5\u5FD7\u5931\u8D25\uFF1A\n{ex.Message}",
                "\u9519\u8BEF", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
