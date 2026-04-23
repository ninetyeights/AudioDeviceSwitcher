using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using NAudio.CoreAudioApi;

namespace AudioDeviceSwitcher;

public record AudioDeviceInfo(string Id, string Name, bool IsDefault, ImageSource? Icon, bool IsDisabled = false);

public static class AudioDeviceService
{
    public static List<AudioDeviceInfo> GetPlaybackDevices(bool includeDisabled = false)
    {
        return GetDevices(DataFlow.Render, includeDisabled);
    }

    public static List<AudioDeviceInfo> GetRecordingDevices(bool includeDisabled = false)
    {
        return GetDevices(DataFlow.Capture, includeDisabled);
    }

    private static List<AudioDeviceInfo> GetDevices(DataFlow dataFlow, bool includeDisabled)
    {
        using var enumerator = new MMDeviceEnumerator();
        MMDevice? defaultDevice = null;
        try
        {
            defaultDevice = enumerator.GetDefaultAudioEndpoint(dataFlow, Role.Multimedia);
        }
        catch (COMException) { }

        var stateMask = includeDisabled
            ? DeviceState.Active | DeviceState.Disabled
            : DeviceState.Active;
        var devices = enumerator.EnumerateAudioEndPoints(dataFlow, stateMask);
        var result = new List<AudioDeviceInfo>();
        foreach (var device in devices)
        {
            ImageSource? icon = null;
            try
            {
                icon = DeviceIconHelper.GetDeviceIcon(device.IconPath);
            }
            catch { }

            result.Add(new AudioDeviceInfo(
                device.ID,
                device.FriendlyName,
                defaultDevice != null && device.ID == defaultDevice.ID,
                icon,
                device.State == DeviceState.Disabled
            ));
        }
        result.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.CurrentCultureIgnoreCase));
        return result;
    }

    private static readonly PropertyKey PKEY_Device_EnumeratorName = new(
        new Guid(0xa45c254e, 0xdf1c, 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), 24);

    public static bool HasBluetoothDevice()
    {
        using var enumerator = new MMDeviceEnumerator();
        var devices = enumerator.EnumerateAudioEndPoints(DataFlow.All, DeviceState.Active);
        foreach (var device in devices)
        {
            try
            {
                var value = device.Properties[PKEY_Device_EnumeratorName].Value;
                if (value is string enumName
                    && enumName.Contains("BTHENUM", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { }
        }
        return false;
    }

    public static void SetDefaultDevice(string deviceId)
    {
        var policy = new PolicyConfigClient();
        policy.SetDefaultEndpoint(deviceId, Role.Multimedia);
        policy.SetDefaultEndpoint(deviceId, Role.Communications);
    }

    public static void SetDeviceEnabled(string deviceId, bool enabled)
    {
        var policy = new PolicyConfigClient();
        var policyConfig = (IPolicyConfig)policy;
        policyConfig.SetEndpointVisibility(deviceId, enabled);
    }

    public static (string? Id, string? Name) GetCommunicationsDefault(DataFlow dataFlow)
    {
        using var enumerator = new MMDeviceEnumerator();
        try
        {
            var dev = enumerator.GetDefaultAudioEndpoint(dataFlow, Role.Communications);
            return (dev.ID, dev.FriendlyName);
        }
        catch (COMException) { return (null, null); }
    }
}

// COM interop for IPolicyConfig (undocumented Windows API for setting default audio device)
[ComImport]
[Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")]
internal class PolicyConfigClient
{
}

[ComImport]
[Guid("f8679f50-850a-41cf-9c72-430f290290c8")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IPolicyConfig
{
    // Cycled through methods to get to SetDefaultEndpoint
    void GetMixFormat();
    void GetDeviceFormat();
    void ResetDeviceFormat();
    void SetDeviceFormat();
    void GetProcessingPeriod();
    void SetProcessingPeriod();
    void GetShareMode();
    void SetShareMode();
    void GetPropertyValue();
    void SetPropertyValue();
    void SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string deviceId, Role role);
    void SetEndpointVisibility([MarshalAs(UnmanagedType.LPWStr)] string deviceId, [MarshalAs(UnmanagedType.Bool)] bool visible);
}

internal static class PolicyConfigExtensions
{
    public static void SetDefaultEndpoint(this PolicyConfigClient client, string deviceId, Role role)
    {
        var policyConfig = (IPolicyConfig)client;
        policyConfig.SetDefaultEndpoint(deviceId, role);
    }
}
