using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml.Linq;

internal static class UiChecks
{
    // Render XAML without invoking app constructors, device changes, timers or settings writes.
    public static void Render(string projectDirectory, string outputDirectory)
    {
        var application = new Application();
        application.Resources.MergedDictionaries.Add(new ResourceDictionary
        {
            Source = new Uri("pack://application:,,,/AudioDeviceSwitcher;component/Theme.xaml")
        });
        Directory.CreateDirectory(outputDirectory);
        foreach (var file in Directory.GetFiles(projectDirectory, "*.xaml"))
        {
            var document = XDocument.Load(file);
            var root = document.Root!;
            if (root.Name.LocalName != "Window") continue;
            root.Attribute(XName.Get("Class", "http://schemas.microsoft.com/winfx/2006/xaml"))?.Remove();
            foreach (var node in root.DescendantsAndSelf())
            {
                var type = typeof(Window).Assembly.GetType("System.Windows.Controls." + node.Name.LocalName)
                    ?? typeof(Window).Assembly.GetType("System.Windows.Controls.Primitives." + node.Name.LocalName)
                    ?? typeof(Window).Assembly.GetType("System.Windows." + node.Name.LocalName);
                foreach (var attribute in node.Attributes().ToList())
                    if (type?.GetEvent(attribute.Name.LocalName) != null) attribute.Remove();
            }
            var markup = document.ToString().Replace("pack://application:,,,/Resources/app.ico", "pack://application:,,,/AudioDeviceSwitcher;component/Resources/app.ico");
            var window = (Window)XamlReader.Parse(markup);
            var content = (FrameworkElement)window.Content;
            if (window.FindName("SessionList") is ItemsControl sessions)
            {
                var devices = new[] { new AudioDeviceSwitcher.AppAudioWindow.DeviceOption(null, "跟随系统"), new AudioDeviceSwitcher.AppAudioWindow.DeviceOption("test", "Headphones (Realtek USB Audio)") };
                sessions.ItemsSource = new[] { new AudioRow { DisplayName = "Google Chrome — 长应用名称显示检查", SubText = @"C:\Program Files\Google\Chrome\Application\chrome.exe", OutputOptions = devices, InputOptions = devices, SelectedOutput = devices[0], SelectedInput = devices[0] } };
            }
            foreach (string name in new[] { "OutputBox", "InputBox", "ProfileChoice", "SessionBox", "AppProfileBox" })
                if (window.FindName(name) is ComboBox combo)
                {
                    combo.DisplayMemberPath = "";
                    combo.ItemsSource = new[] { "工作耳机 / Headphones (Realtek USB Audio)", "跟随系统" };
                    combo.SelectedIndex = 0;
                }
            // Give empty lists realistic rows, including deliberately long device names.
            if (window.FindName("ScheduleList") is DataGrid schedules)
                schedules.ItemsSource = new[] { new { State = "启用", Profile = "工作耳机", Timing = "09:00 · 周一、周二、周三", Next = "09-22 09:00", Result = "等待下一次执行" }, new { State = "停用", Profile = "夜间娱乐", Timing = "18:00 · 每天", Next = "—", Result = "已跳过：配置已锁定，请先解锁" } };
            if (window.FindName("ProfileList") is ItemsControl profiles && Path.GetFileName(file) == "AppProfilesWindow.xaml")
                profiles.ItemsSource = new[] { new { Name = "会议与录音", OutputDeviceName = "Headphones (Realtek USB Audio) — 很长的设备名称", InputDeviceName = "Microphone (Realtek USB Audio)", Id = Guid.Empty } };
            foreach (string name in new[] { "NameBox", "TimeInput", "NicknameBox" })
                if (window.FindName(name) is TextBox input && name != "TimeInput") input.Text = "工作耳机";
            if (window.FindName("OriginalNameBlock") is TextBlock original) original.Text = "Headphones (Realtek USB Audio)";
            double width = window.Width;
            double height = window.SizeToContent is SizeToContent.Height or SizeToContent.WidthAndHeight ? double.PositiveInfinity : window.Height - 32;
            content.Measure(new Size(width, height));
            if (double.IsInfinity(height)) height = content.DesiredSize.Height;
            content.Arrange(new Rect(0, 0, width, height));
            content.UpdateLayout();
            var bitmap = new RenderTargetBitmap((int)Math.Ceiling(width), (int)Math.Ceiling(height), 96, 96, PixelFormats.Pbgra32);
            var white = new DrawingVisual();
            using (var drawing = white.RenderOpen()) drawing.DrawRectangle(Brushes.White, null, new Rect(0, 0, width, height));
            bitmap.Render(white);
            bitmap.Render(content);
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bitmap));
            using (var stream = File.Create(Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(file) + ".png"))) encoder.Save(stream);
            window.Close();
            Console.WriteLine("PASS render " + Path.GetFileName(file));
        }
        application.Shutdown();
    }

    public class AudioRow
    {
        public string DisplayName { get; set; } = "";
        public string SubText { get; set; } = "";
        public string RouteStatusText => "部分路由等待音频会话";
        public bool HasOutputSession => true;
        public bool IsMuted { get; set; }
        public double Volume { get; set; } = 0.6;
        public string VolumePercent => "60%";
        public bool IsDrifted => false;
        public bool IsOutputDrifted => false;
        public bool IsInputDrifted => false;
        public object? Icon => null;
        public AudioDeviceSwitcher.AppAudioWindow.DeviceOption[] OutputOptions { get; set; } = [];
        public AudioDeviceSwitcher.AppAudioWindow.DeviceOption[] InputOptions { get; set; } = [];
        public AudioDeviceSwitcher.AppAudioWindow.DeviceOption? SelectedOutput { get; set; }
        public AudioDeviceSwitcher.AppAudioWindow.DeviceOption? SelectedInput { get; set; }
    }
}
