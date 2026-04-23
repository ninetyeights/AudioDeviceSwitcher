using System.Runtime.InteropServices;
using Microsoft.Win32;
using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;

namespace AudioDeviceSwitcher;

// Win11-friendly notification service using ToastNotificationManager with Tag/Group.
// New toasts with the same Tag automatically replace earlier ones in the notification center,
// so rapid profile switches don't produce stacked stale notifications.
public static class ToastService
{
    public const string Aumid = "AudioDeviceSwitcher.App";
    private const string Group = "ADS";
    private const string AppDisplayName = "音频切换助手";

    // Tags — each distinct kind of notification uses its own Tag so they don't overwrite each other.
    public const string TagProfileSwitch = "profile-switch";
    public const string TagDeviceChange = "device-change";
    public const string TagBluetooth = "bluetooth";
    public const string TagAppDrift = "app-drift";

    // Call once at startup — sets the process AUMID (so notifications display with the right
    // identity) and writes registry metadata (display name + icon) for the notification center.
    public static void Initialize()
    {
        try { SetCurrentProcessExplicitAppUserModelID(Aumid); } catch { }
        try { RegisterAppUserModelId(); } catch { }
    }

    public static void Show(string tag, string title, string body)
    {
        try
        {
            var xml = ToastNotificationManager.GetTemplateContent(ToastTemplateType.ToastText02);
            var texts = xml.GetElementsByTagName("text");
            texts[0].AppendChild(xml.CreateTextNode(title));
            texts[1].AppendChild(xml.CreateTextNode(body));

            var toast = new ToastNotification(xml)
            {
                Tag = tag,
                Group = Group,
                ExpirationTime = DateTimeOffset.Now.AddSeconds(5),
            };

            ToastNotificationManager.CreateToastNotifier(Aumid).Show(toast);
        }
        catch { /* toast may fail on some SKUs / unregistered AUMIDs — silently ignore */ }
    }

    public static void ClearAll()
    {
        try { ToastNotificationManager.History.Clear(Aumid); } catch { }
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern void SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string AppID);

    private static void RegisterAppUserModelId()
    {
        // HKCU\Software\Classes\AppUserModelId\<AUMID>\(DisplayName, IconUri)
        using var key = Registry.CurrentUser.CreateSubKey($@"Software\Classes\AppUserModelId\{Aumid}");
        if (key == null) return;
        key.SetValue("DisplayName", AppDisplayName, RegistryValueKind.String);
        var exePath = Environment.ProcessPath;
        if (!string.IsNullOrEmpty(exePath))
            key.SetValue("IconUri", exePath, RegistryValueKind.String);
    }
}
