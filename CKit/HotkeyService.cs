using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace AudioDeviceSwitcher;

public partial class HotkeyService : IDisposable
{
    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool UnregisterHotKey(IntPtr hWnd, int id);

    private const int WM_HOTKEY = 0x0312;
    private readonly Window _window;
    private readonly Dictionary<int, Action> _callbacks = new();
    private int _nextId = 1;
    private HwndSource? _source;

    public HotkeyService(Window window)
    {
        _window = window;
        var helper = new WindowInteropHelper(_window);
        if (helper.Handle != IntPtr.Zero)
        {
            _source = HwndSource.FromHwnd(helper.Handle);
            _source?.AddHook(WndProc);
        }
        else
        {
            _window.SourceInitialized += (_, _) =>
            {
                _source = HwndSource.FromHwnd(new WindowInteropHelper(_window).Handle);
                _source?.AddHook(WndProc);
            };
        }
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WM_HOTKEY && _callbacks.TryGetValue((int)wParam, out var callback))
        {
            callback();
            handled = true;
        }
        return IntPtr.Zero;
    }

    public int Register(int modifiers, int key, Action callback)
    {
        var handle = new WindowInteropHelper(_window).Handle;
        if (handle == IntPtr.Zero) return -1;

        var id = _nextId++;
        if (RegisterHotKey(handle, id, modifiers, key))
        {
            _callbacks[id] = callback;
            return id;
        }
        return -1;
    }

    public void Unregister(int id)
    {
        if (id <= 0) return;
        var handle = new WindowInteropHelper(_window).Handle;
        if (handle != IntPtr.Zero)
            UnregisterHotKey(handle, id);
        _callbacks.Remove(id);
    }

    public void UnregisterAll()
    {
        var handle = new WindowInteropHelper(_window).Handle;
        foreach (var id in _callbacks.Keys)
        {
            if (handle != IntPtr.Zero)
                UnregisterHotKey(handle, id);
        }
        _callbacks.Clear();
    }

    public void Dispose()
    {
        UnregisterAll();
        _source?.RemoveHook(WndProc);
    }
}
