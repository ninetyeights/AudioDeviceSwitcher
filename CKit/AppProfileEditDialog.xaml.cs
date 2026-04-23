using System.Windows;

namespace AudioDeviceSwitcher;

public partial class AppProfileEditDialog : Window
{
    public record DeviceOption(string? Id, string Name);

    private static readonly DeviceOption NoneOption = new(null, "跟随系统");

    public AppProfile Result { get; private set; }

    public AppProfileEditDialog(AppProfile? existing = null)
    {
        InitializeComponent();

        Result = existing ?? new AppProfile();
        NameBox.Text = Result.Name;

        var outputs = new List<DeviceOption> { NoneOption };
        foreach (var d in AudioDeviceService.GetPlaybackDevices())
            outputs.Add(new DeviceOption(d.Id, d.Name));
        OutputBox.ItemsSource = outputs;
        OutputBox.SelectedItem = outputs.FirstOrDefault(o =>
            string.Equals(o.Id, Result.OutputDeviceId, StringComparison.OrdinalIgnoreCase)) ?? NoneOption;

        var inputs = new List<DeviceOption> { NoneOption };
        foreach (var d in AudioDeviceService.GetRecordingDevices())
            inputs.Add(new DeviceOption(d.Id, d.Name));
        InputBox.ItemsSource = inputs;
        InputBox.SelectedItem = inputs.FirstOrDefault(o =>
            string.Equals(o.Id, Result.InputDeviceId, StringComparison.OrdinalIgnoreCase)) ?? NoneOption;

        Loaded += (_, _) => { NameBox.Focus(); NameBox.SelectAll(); };
    }

    private void OK_Click(object sender, RoutedEventArgs e)
    {
        var name = NameBox.Text.Trim();
        if (string.IsNullOrEmpty(name))
        {
            MessageBox.Show("请输入配置名称。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var output = OutputBox.SelectedItem as DeviceOption ?? NoneOption;
        var input = InputBox.SelectedItem as DeviceOption ?? NoneOption;

        Result = Result with
        {
            Name = name,
            OutputDeviceId = output.Id,
            OutputDeviceName = output.Id == null ? null : output.Name,
            InputDeviceId = input.Id,
            InputDeviceName = input.Id == null ? null : input.Name,
        };

        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
}
