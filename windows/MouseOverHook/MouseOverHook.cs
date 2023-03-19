using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

class MouseOverHook
{
    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    private static extern IntPtr WindowFromPoint(POINT Point);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

    private const int WH_MOUSE_LL = 14;
    private const int WM_MOUSEMOVE = 0x0200;

    private static LowLevelMouseProc _proc = HookCallback;
    private static IntPtr _hookID = IntPtr.Zero;

    public static void Main()
    {
        _hookID = SetHook(_proc);
        Application.Run();
        UnhookWindowsHookEx(_hookID);
    }

    private static IntPtr SetHook(LowLevelMouseProc proc)
    {
        using (Process curProcess = Process.GetCurrentProcess())
        using (ProcessModule curModule = curProcess.MainModule)
        {
            return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
        }
    }

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_MOUSEMOVE)
        {
            POINT cursorPos;
            GetCursorPos(out cursorPos);
            IntPtr hWnd = WindowFromPoint(cursorPos);

            if (hWnd != IntPtr.Zero)
            {
                uint processID;
                GetWindowThreadProcessId(hWnd, out processID);

                if (processID != 0)
                {
                    string processName = GetProcessName(processID);
                    Console.WriteLine("Mouse-over event detected. Process ID: {0}, Software Name: {1}", processID, processName);
                }
            }
        }

        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    private static string GetProcessName(uint processID)
    {
        try
        {
            Process process = Process.GetProcessById((int)processID);
            return process.ProcessName;
        }
        catch (Exception)
        {
            return "unknown";
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int x;
        public int y;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);
}
