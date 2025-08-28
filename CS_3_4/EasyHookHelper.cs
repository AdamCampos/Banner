using System;
using System.Runtime.InteropServices;
using EasyHook;

namespace ComHookLib.Hooking
{
    internal static class EasyHookHelper
    {
        public static LocalHook CreateHook<TDelegate>(string module, string procName, TDelegate hookHandler, object callback)
            where TDelegate : Delegate
        {
            if (hookHandler == null) throw new ArgumentNullException(nameof(hookHandler));

            IntPtr addr = LocalHook.GetProcAddress(module, procName);
            if (addr == IntPtr.Zero)
                throw new InvalidOperationException($"GetProcAddress falhou: {module}!{procName}");

            var hook = LocalHook.Create(addr, hookHandler, callback);

            // Permitir em todos os threads (nenhum excluído)
            hook.ThreadACL.SetExclusiveACL(Array.Empty<int>());

            return hook;
        }

        public static TDelegate GetOriginalDelegate<TDelegate>(string module, string procName)
            where TDelegate : Delegate
        {
            IntPtr addr = LocalHook.GetProcAddress(module, procName);
            if (addr == IntPtr.Zero)
                throw new InvalidOperationException($"GetProcAddress falhou: {module}!{procName}");

            return (TDelegate)Marshal.GetDelegateForFunctionPointer(addr, typeof(TDelegate));
        }
    }
}
