using System.Windows;

namespace AudioDeviceSwitcher;

public partial class DeviceRenameDialog : Window
{
    public string? Nickname { get; private set; }
    public bool Cleared { get; private set; }

    public DeviceRenameDialog(string originalName, string? currentNickname)
    {
        InitializeComponent();
        OriginalNameBlock.Text = $"原始名称: {originalName}";
        NicknameBox.Text = currentNickname ?? "";
        Loaded += (_, _) => { NicknameBox.Focus(); NicknameBox.SelectAll(); };
    }

    private void OK_Click(object sender, RoutedEventArgs e)
    {
        var v = NicknameBox.Text.Trim();
        Nickname = string.IsNullOrEmpty(v) ? null : v;
        Cleared = Nickname == null;
        DialogResult = true;
    }

    private void Clear_Click(object sender, RoutedEventArgs e)
    {
        Nickname = null;
        Cleared = true;
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
}
