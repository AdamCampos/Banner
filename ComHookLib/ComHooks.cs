using System;
using System.Runtime.InteropServices;
using System.Threading;
using ComHookLib;
using EasyHook;

namespace ComHook
{
    /// <summary>
    /// Instala hooks em APIs COM básicas e registra chamadas.
    /// Compatível com C# 7.3 (sem target-typed new, sem using declarations).
    /// </summary>
    public static class ComHooks
    {
        // Delegates das funções hookadas
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int CoCreateInstance_Delegate(ref Guid rclsid, IntPtr pUnkOuter, uint dwClsContext, ref Guid riid, out IntPtr ppv);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int CoGetClassObject_Delegate(ref Guid rclsid, uint dwClsContext, IntPtr pServerInfo, ref Guid riid, out IntPtr ppv);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int CoCreateInstanceEx_Delegate(ref Guid rclsid, IntPtr pUnkOuter, uint dwClsCtx, IntPtr pServerInfo, uint dwCount, IntPtr pResults);

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
        private delegate int CoGetObject_Delegate(string pszName, IntPtr pBindingOptions, ref Guid riid, out IntPtr ppv);

        // Original pointers
        private static CoCreateInstance_Delegate _origCreate;
        private static CoGetClassObject_Delegate _origGetClassObject;
        private static CoCreateInstanceEx_Delegate _origCreateEx;
        private static CoGetObject_Delegate _origGetObject;

        // Hooks
        private static LocalHook _hookCreate;
        private static LocalHook _hookGetClassObject;
        private static LocalHook _hookCreateEx;
        private static LocalHook _hookGetObject;

        private static volatile bool _installed;

        // Consts de contexto
        private const uint CLSCTX_INPROC_SERVER = 0x1;
        private const uint CLSCTX_INPROC_HANDLER = 0x2;
        private const uint CLSCTX_LOCAL_SERVER = 0x4;
        private const uint CLSCTX_REMOTE_SERVER = 0x10;

        // IID comuns
        private static readonly Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        // -------------------- API wrappers --------------------

        private static int CoCreateInstance_Hook(ref Guid rclsid, IntPtr pUnkOuter, uint ctx, ref Guid riid, out IntPtr ppv)
        {
            int hr = 0;
            ppv = IntPtr.Zero;
            try
            {
                hr = _origCreate(ref rclsid, pUnkOuter, ctx, ref riid, out ppv);

                string clsid = rclsid.ToString();
                string iid = riid.ToString();
                uint uhr = (uint)hr;
                string extra = ppv == IntPtr.Zero ? "ppv=null" : "ppv=ok";

                ComLogger.Csv("CoCreateInstance", clsid, iid, (int)ctx, uhr, ThreadHelper.ThreadId(), "", extra, null);

                // Tentativa básica de descobrir módulo carregado (não obrigatório)
                string modulePath = ModuleHelper.TryGetServerModulePath(rclsid, ctx);
                if (!string.IsNullOrEmpty(modulePath))
                {
                    // chamada com nomeado modulePath para cobrir os pontos que usam named arg
                    ComLogger.Csv("CoCreateInstance", clsid, iid, (int)ctx, uhr, modulePath);
                }

                // Se a interface suportar IDispatch/IConnectionPointContainer, apenas logamos — sem Advises automáticos aqui
                return hr;
            }
            catch (Exception ex)
            {
                ComLogger.Err("CoCreateInstance_Hook", ex);
                return hr != 0 ? hr : Marshal.GetHRForException(ex);
            }
        }

        private static int CoGetClassObject_Hook(ref Guid rclsid, uint ctx, IntPtr pServerInfo, ref Guid riid, out IntPtr ppv)
        {
            int hr = 0;
            ppv = IntPtr.Zero;
            try
            {
                hr = _origGetClassObject(ref rclsid, ctx, pServerInfo, ref riid, out ppv);

                string clsid = rclsid.ToString();
                string iid = riid.ToString();
                uint uhr = (uint)hr;
                string extra = ppv == IntPtr.Zero ? "ppv=null" : "ppv=ok";

                ComLogger.Csv("CoGetClassObject", clsid, iid, (int)ctx, uhr, ThreadHelper.ThreadId(), "", extra, null);
                return hr;
            }
            catch (Exception ex)
            {
                ComLogger.Err("CoGetClassObject_Hook", ex);
                return hr != 0 ? hr : Marshal.GetHRForException(ex);
            }
        }

        private static int CoCreateInstanceEx_Hook(ref Guid rclsid, IntPtr pUnkOuter, uint ctx, IntPtr pServerInfo, uint dwCount, IntPtr pResults)
        {
            int hr = 0;
            try
            {
                hr = _origCreateEx(ref rclsid, pUnkOuter, ctx, pServerInfo, dwCount, pResults);

                string clsid = rclsid.ToString();
                uint uhr = (uint)hr;

                ComLogger.Csv("CoCreateInstanceEx", clsid, IID_IUnknown.ToString(), (int)ctx, uhr, ThreadHelper.ThreadId(), "", "results=" + dwCount, null);

                string modulePath = ModuleHelper.TryGetServerModulePath(rclsid, ctx);
                if (!string.IsNullOrEmpty(modulePath))
                {
                    ComLogger.Csv("CoCreateInstanceEx", clsid, IID_IUnknown.ToString(), (int)ctx, uhr, modulePath);
                }

                return hr;
            }
            catch (Exception ex)
            {
                ComLogger.Err("CoCreateInstanceEx_Hook", ex);
                return hr != 0 ? hr : Marshal.GetHRForException(ex);
            }
        }

        private static int CoGetObject_Hook(string pszName, IntPtr pBindingOptions, ref Guid riid, out IntPtr ppv)
        {
            int hr = 0;
            ppv = IntPtr.Zero;
            try
            {
                hr = _origGetObject(pszName, pBindingOptions, ref riid, out ppv);

                string iid = riid.ToString();
                uint uhr = (uint)hr;
                string extra = ppv == IntPtr.Zero ? "ppv=null" : "ppv=ok";

                ComLogger.Csv("CoGetObject", pszName ?? "", iid, 0, uhr, ThreadHelper.ThreadId(), "", extra, null);
                return hr;
            }
            catch (Exception ex)
            {
                ComLogger.Err("CoGetObject_Hook", ex);
                return hr != 0 ? hr : Marshal.GetHRForException(ex);
            }
        }

        // -------------------- Instalação / Remoção --------------------

        public static void Install()
        {
            if (_installed)
                return;

            // Resolve endereços
            IntPtr pCreate = LocalHook.GetProcAddress("ole32.dll", "CoCreateInstance");
            IntPtr pGetClass = LocalHook.GetProcAddress("ole32.dll", "CoGetClassObject");
            IntPtr pCreateEx = LocalHook.GetProcAddress("ole32.dll", "CoCreateInstanceEx");
            IntPtr pGetObj = LocalHook.GetProcAddress("ole32.dll", "CoGetObject");

            _origCreate = (CoCreateInstance_Delegate)Marshal.GetDelegateForFunctionPointer(pCreate, typeof(CoCreateInstance_Delegate));
            _origGetClassObject = (CoGetClassObject_Delegate)Marshal.GetDelegateForFunctionPointer(pGetClass, typeof(CoGetClassObject_Delegate));
            _origCreateEx = (CoCreateInstanceEx_Delegate)Marshal.GetDelegateForFunctionPointer(pCreateEx, typeof(CoCreateInstanceEx_Delegate));
            _origGetObject = (CoGetObject_Delegate)Marshal.GetDelegateForFunctionPointer(pGetObj, typeof(CoGetObject_Delegate));

            // Cria hooks
            _hookCreate = LocalHook.Create(pCreate, new CoCreateInstance_Delegate(CoCreateInstance_Hook), null);
            _hookGetClassObject = LocalHook.Create(pGetClass, new CoGetClassObject_Delegate(CoGetClassObject_Hook), null);
            _hookCreateEx = LocalHook.Create(pCreateEx, new CoCreateInstanceEx_Delegate(CoCreateInstanceEx_Hook), null);
            _hookGetObject = LocalHook.Create(pGetObj, new CoGetObject_Delegate(CoGetObject_Hook), null);

            // Libera para todos os threads (exclusivo thread 0)
            _hookCreate.ThreadACL.SetExclusiveACL(new[] { 0 });
            _hookGetClassObject.ThreadACL.SetExclusiveACL(new[] { 0 });
            _hookCreateEx.ThreadACL.SetExclusiveACL(new[] { 0 });
            _hookGetObject.ThreadACL.SetExclusiveACL(new[] { 0 });

            _installed = true;
            ComLogger.Write("ComHooks installed.");
        }

        /// <summary>
        /// Desinstala/remover hooks com segurança. (NOVO)
        /// </summary>
        public static void Uninstall()
        {
            try
            {
                if (_hookCreate != null) { _hookCreate.ThreadACL.SetInclusiveACL(new[] { 0 }); _hookCreate.Dispose(); _hookCreate = null; }
                if (_hookGetClassObject != null) { _hookGetClassObject.ThreadACL.SetInclusiveACL(new[] { 0 }); _hookGetClassObject.Dispose(); _hookGetClassObject = null; }
                if (_hookCreateEx != null) { _hookCreateEx.ThreadACL.SetInclusiveACL(new[] { 0 }); _hookCreateEx.Dispose(); _hookCreateEx = null; }
                if (_hookGetObject != null) { _hookGetObject.ThreadACL.SetInclusiveACL(new[] { 0 }); _hookGetObject.Dispose(); _hookGetObject = null; }
            }
            catch (Exception ex)
            {
                ComLogger.Err("ComHooks.Uninstall", ex);
            }
            finally
            {
                _installed = false;
            }
        }

        // -------------------- Helpers --------------------

        private static class ThreadHelper
        {
            public static int ThreadId()
            {
                try { return Thread.CurrentThread.ManagedThreadId; }
                catch { return 0; }
            }
        }

        private static class ModuleHelper
        {
            // Tenta localizar caminho do servidor COM pelo CLSID e ctx.
            // Implementação simplificada para não depender de registry intensivo: retorna string vazia se não achar.
            public static string TryGetServerModulePath(Guid clsid, uint ctx)
            {
                try
                {
                    // Para INPROC_SERVER, normalmente está em HKCR\CLSID\{clsid}\InprocServer32
                    // Aqui deixamos como "não resolvido" para manter compatibilidade sem travar.
                    // Você pode implementar resolução por registry se desejar.
                    return string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }
        }
    }
}
