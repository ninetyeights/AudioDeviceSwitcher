using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace AudioDeviceSwitcher;

internal static partial class DeviceIconHelper
{
    [LibraryImport("user32.dll", EntryPoint = "PrivateExtractIconsW", StringMarshalling = StringMarshalling.Utf16)]
    private static partial int PrivateExtractIcons(
        string szFileName, int nIconIndex, int cxIcon, int cyIcon,
        nint[]? phicon, int[]? piconid, int nIcons, int flags);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool DestroyIcon(nint hIcon);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct SHFILEINFO
    {
        public nint hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }

    private const uint SHGFI_ICON = 0x100;
    private const uint SHGFI_LARGEICON = 0x0;
    private const uint SHGFI_SMALLICON = 0x1;

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern nint SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

    // Frozen BitmapSources are immutable and safe to share across the app; caching them
    // avoids re-extracting icons on every device/session enumeration (called every few
    // seconds by timers). Null results are cached too — failed extractions stay failed.
    private static readonly ConcurrentDictionary<string, ImageSource?> _cache = new();

    public static ImageSource? GetExeIcon(string exePath, int size = 32)
    {
        if (string.IsNullOrEmpty(exePath)) return null;
        var key = $"exe|{size}|{exePath}";
        return _cache.GetOrAdd(key, _ => ExtractIcon(exePath, 0, size) ?? ExtractIconViaShell(exePath, size));
    }

    // PrivateExtractIcons only finds icons embedded as Win32 resources; it comes up empty
    // for packaged (UWP/MSIX) apps and some protected exes whose icon is resolved by the
    // shell through other means (app manifest, association). SHGetFileInfo goes through the
    // same resolution path Explorer/Volume Mixer use, so it succeeds where the direct
    // extraction doesn't.
    private static ImageSource? ExtractIconViaShell(string filePath, int size)
    {
        var shfi = new SHFILEINFO();
        var flags = SHGFI_ICON | (size <= 16 ? SHGFI_SMALLICON : SHGFI_LARGEICON);
        var result = SHGetFileInfo(filePath, 0, ref shfi, (uint)Marshal.SizeOf<SHFILEINFO>(), flags);
        if (result == 0 || shfi.hIcon == 0) return null;
        try
        {
            var source = Imaging.CreateBitmapSourceFromHIcon(
                shfi.hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            source.Freeze();
            return source;
        }
        catch { return null; }
        finally { DestroyIcon(shfi.hIcon); }
    }

    public static ImageSource? GetDeviceIcon(string iconPath, int size = 48)
    {
        if (string.IsNullOrEmpty(iconPath)) return null;

        // Icon path format: "%SystemRoot%\System32\ddores.dll,-2033"
        var expanded = Environment.ExpandEnvironmentVariables(iconPath);
        var parts = expanded.Split(',');
        if (parts.Length < 2) return null;

        var filePath = parts[0].Trim();
        if (!int.TryParse(parts[1].Trim(), out var iconIndex)) return null;

        var key = $"dev|{size}|{filePath}|{iconIndex}";
        return _cache.GetOrAdd(key, _ => ExtractIcon(filePath, iconIndex, size));
    }

    private static ImageSource? ExtractIcon(string filePath, int iconIndex, int size)
    {
        var hIcons = new nint[1];
        var ids = new int[1];
        var count = PrivateExtractIcons(filePath, iconIndex, size, size, hIcons, ids, 1, 0);
        if (count == 0 || hIcons[0] == 0) return null;
        try
        {
            var source = Imaging.CreateBitmapSourceFromHIcon(
                hIcons[0], Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            source.Freeze();
            return source;
        }
        catch { return null; }
        finally { DestroyIcon(hIcons[0]); }
    }
}
