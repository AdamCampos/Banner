using System;
using System.Runtime.InteropServices;
using System.Text;

namespace ComHookLib
{
    [Flags]
    internal enum CLSCTX : uint
    {
        INPROC_SERVER = 0x1,
        INPROC_HANDLER = 0x2,
        LOCAL_SERVER = 0x4,
        INPROC_SERVER16 = 0x8,
        REMOTE_SERVER = 0x10,
        INPROC_HANDLER16 = 0x20,
        NO_CODE_DOWNLOAD = 0x400,
        NO_CUSTOM_MARSHAL = 0x1000,
        ENABLE_CODE_DOWNLOAD = 0x2000,
        NO_FAILURE_LOG = 0x4000,
        DISABLE_AAA = 0x8000,
        ENABLE_AAA = 0x10000,
        FROM_DEFAULT_CONTEXT = 0x20000,
        ACTIVATE_32_BIT_SERVER = 0x40000,
        ACTIVATE_64_BIT_SERVER = 0x80000,
        ENABLE_CLOAKING = 0x100000,
        APPCONTAINER = 0x400000,
        ACTIVATE_AAA_AS_IU = 0x800000,
        PS_DLL = 0x80000000
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    internal struct COSERVERINFO
    {
        public uint dwReserved1;
        public string pwszName;
        public IntPtr pAuthInfo;
        public uint dwReserved2;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MULTI_QI
    {
        public IntPtr pIID;   // REFIID*
        public IntPtr pItf;   // IUnknown*
        public int hr;        // HRESULT
    }

    internal static class Native
    {
        [DllImport("kernel32.dll")]
        public static extern int GetCurrentThreadId();

        internal const uint EVENT_OBJECT_CREATE = 0x8000;
        internal const uint EVENT_OBJECT_SHOW = 0x8002;
        internal const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
        internal const uint WINEVENT_OUTOFCONTEXT = 0x0000;

        internal delegate void WinEventDelegate(
            IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        internal static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc,
            WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        internal static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassNameW(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        internal static string GetClassNameSafe(IntPtr hwnd)
        {
            var sb = new StringBuilder(256);
            try { GetClassNameW(hwnd, sb, sb.Capacity); } catch { }
            return sb.ToString();
        }

        internal static string GetWindowTextSafe(IntPtr hwnd)
        {
            var sb = new StringBuilder(1024);
            try { GetWindowTextW(hwnd, sb, sb.Capacity); } catch { }
            return sb.ToString();
        }

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int ProgIDFromCLSID(ref Guid clsid, out IntPtr lplpszProgID);

        internal static string ProgIdFromClsidSafe(Guid clsid)
        {
            IntPtr ptr = IntPtr.Zero;
            try
            {
                int hr = ProgIDFromCLSID(ref clsid, out ptr);
                if (hr != 0 || ptr == IntPtr.Zero)
                    return null;
                string s = Marshal.PtrToStringUni(ptr);
                Marshal.FreeCoTaskMem(ptr);
                ptr = IntPtr.Zero;
                return s;
            }
            catch
            {
                try { if (ptr != IntPtr.Zero) Marshal.FreeCoTaskMem(ptr); } catch { }
                return null;
            }
        }
    }
}
