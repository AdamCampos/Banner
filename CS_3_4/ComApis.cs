using System;
using System.Runtime.InteropServices;

namespace ComHookLib.Hooking
{
    // Delegates que refletem as assinaturas nativas de OLE32 (stdcall)
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    internal delegate int CoCreateInstance_Delegate(
        ref Guid rclsid,
        IntPtr pUnkOuter,
        ComHookLib.CLSCTX dwClsContext,
        ref Guid riid,
        out IntPtr ppv
    );

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    internal delegate int CoGetClassObject_Delegate(
        ref Guid rclsid,
        ComHookLib.CLSCTX dwClsContext,
        IntPtr pServerInfo,   // COSERVERINFO*
        ref Guid riid,
        out IntPtr ppv
    );

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    internal delegate int CoCreateInstanceEx_Delegate(
        ref Guid rclsid,
        IntPtr pUnkOuter,
        ComHookLib.CLSCTX dwClsCtx,
        IntPtr pServerInfo,    // COSERVERINFO*
        uint dwCount,
        IntPtr pResults        // MULTI_QI*
    );
}
