using System.Runtime.InteropServices;
using NAudio.CoreAudioApi;

namespace AudioDeviceSwitcher;

// COM interop for IAudioPolicyConfig (undocumented WinRT API used by Windows Settings'
// per-app audio device override in Volume Mixer). Reverse-engineered by EarTrumpet et al.
// IID ab3d4648-... applies to Windows 10 1809+ and Windows 11.
public static class AppAudioRoutingService
{
    private const string RuntimeClassName = "Windows.Media.Internal.AudioPolicyConfig";

    // Device ID must be wrapped as a device-interface path for IAudioPolicyConfig.
    // Source: EarTrumpet AudioPolicyConfigService.cs.
    private const string MMDEVAPI_TOKEN = @"\\?\SWD#MMDEVAPI#";
    private const string DEVINTERFACE_AUDIO_RENDER = "#{e6327cad-dcec-4949-ae8a-991e976a79d2}";
    private const string DEVINTERFACE_AUDIO_CAPTURE = "#{2eef81be-33fa-4800-9670-1cd474972c3f}";

    private static string WrapDeviceId(string deviceId, DataFlow flow) =>
        $"{MMDEVAPI_TOKEN}{deviceId}{(flow == DataFlow.Render ? DEVINTERFACE_AUDIO_RENDER : DEVINTERFACE_AUDIO_CAPTURE)}";

    private static string UnwrapDeviceId(string wrapped)
    {
        if (wrapped.StartsWith(MMDEVAPI_TOKEN)) wrapped = wrapped[MMDEVAPI_TOKEN.Length..];
        if (wrapped.EndsWith(DEVINTERFACE_AUDIO_RENDER)) wrapped = wrapped[..^DEVINTERFACE_AUDIO_RENDER.Length];
        if (wrapped.EndsWith(DEVINTERFACE_AUDIO_CAPTURE)) wrapped = wrapped[..^DEVINTERFACE_AUDIO_CAPTURE.Length];
        return wrapped;
    }

    [DllImport("combase.dll", PreserveSig = false)]
    private static extern void WindowsCreateString(
        [MarshalAs(UnmanagedType.LPWStr)] string sourceString,
        int length,
        out IntPtr hstring);

    [DllImport("combase.dll", PreserveSig = false)]
    private static extern void WindowsDeleteString(IntPtr hstring);

    [DllImport("combase.dll")]
    private static extern IntPtr WindowsGetStringRawBuffer(IntPtr hstring, out uint length);

    [DllImport("combase.dll", PreserveSig = false)]
    private static extern void RoGetActivationFactory(
        IntPtr activatableClassId,
        [In] ref Guid iid,
        out IntPtr factory);

    [ComImport]
    [Guid("ab3d4648-e242-459f-b02f-541c70306324")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IAudioPolicyConfigFactory
    {
        // Padding — 3 IInspectable slots (GetIids/GetRuntimeClassName/GetTrustLevel)
        // + 19 unused AudioPolicyConfig slots (Win 10 1809+ / Win 11 layout per EarTrumpet)
        int Pad00(); int Pad01(); int Pad02();
        int Pad03(); int Pad04(); int Pad05(); int Pad06();
        int Pad07(); int Pad08(); int Pad09(); int Pad10();
        int Pad11(); int Pad12(); int Pad13(); int Pad14();
        int Pad15(); int Pad16(); int Pad17(); int Pad18();
        int Pad19(); int Pad20(); int Pad21();

        [PreserveSig]
        int SetPersistedDefaultAudioEndpoint(
            uint processId, DataFlow flow, Role role, IntPtr deviceIdHString);

        [PreserveSig]
        int GetPersistedDefaultAudioEndpoint(
            uint processId, DataFlow flow, Role role, out IntPtr deviceIdHString);

        [PreserveSig]
        int ClearAllPersistedApplicationDefaultEndpoints();
    }

    private static IAudioPolicyConfigFactory? _factory;

    private static IAudioPolicyConfigFactory GetFactory()
    {
        if (_factory != null) return _factory;

        WindowsCreateString(RuntimeClassName, RuntimeClassName.Length, out var classId);
        IntPtr pFactory = IntPtr.Zero;
        try
        {
            var iid = typeof(IAudioPolicyConfigFactory).GUID;
            RoGetActivationFactory(classId, ref iid, out pFactory);
            _factory = (IAudioPolicyConfigFactory)Marshal.GetObjectForIUnknown(pFactory);
            return _factory;
        }
        finally
        {
            if (pFactory != IntPtr.Zero) Marshal.Release(pFactory);
            WindowsDeleteString(classId);
        }
    }

    // Set per-app default endpoint. Pass null deviceId to clear (revert to system default).
    // Applies to both Multimedia and Console roles (matches Windows Settings' Volume Mixer).
    public static void SetAppEndpoint(uint processId, DataFlow flow, string? deviceId)
    {
        var factory = GetFactory();
        IntPtr hstr = IntPtr.Zero;
        try
        {
            if (!string.IsNullOrEmpty(deviceId))
            {
                var wrapped = WrapDeviceId(deviceId, flow);
                WindowsCreateString(wrapped, wrapped.Length, out hstr);
            }

            int hr1 = factory.SetPersistedDefaultAudioEndpoint(processId, flow, Role.Multimedia, hstr);
            int hr2 = factory.SetPersistedDefaultAudioEndpoint(processId, flow, Role.Console, hstr);
            if (hr1 < 0) Marshal.ThrowExceptionForHR(hr1);
            if (hr2 < 0) Marshal.ThrowExceptionForHR(hr2);
        }
        finally
        {
            if (hstr != IntPtr.Zero) WindowsDeleteString(hstr);
        }
    }

    public static string? GetAppEndpoint(uint processId, DataFlow flow)
    {
        var factory = GetFactory();
        int hr = factory.GetPersistedDefaultAudioEndpoint(
            processId, flow, Role.Multimedia, out var hstr);
        if (hr < 0 || hstr == IntPtr.Zero) return null;
        try
        {
            var buf = WindowsGetStringRawBuffer(hstr, out var len);
            if (len == 0) return null;
            var wrapped = Marshal.PtrToStringUni(buf, (int)len);
            return wrapped == null ? null : UnwrapDeviceId(wrapped);
        }
        finally
        {
            WindowsDeleteString(hstr);
        }
    }

    public static void ClearAll()
    {
        var factory = GetFactory();
        int hr = factory.ClearAllPersistedApplicationDefaultEndpoints();
        if (hr < 0) Marshal.ThrowExceptionForHR(hr);
    }
}
