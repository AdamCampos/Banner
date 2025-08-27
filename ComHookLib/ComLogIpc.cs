using System;

namespace ComHookLib
{
    public class ComLogIpc : MarshalByRefObject
    {
        public void WriteLine(string message)
        {
            ComLogger.Write("[IPC] " + message);
        }

        public void Csv(string eventType, string clsid, string riid, int hResult,
                        uint processId, int threadId, string modulePath, string serverResolved)
        {
            ComLogger.Csv(eventType, clsid, riid, hResult, processId, threadId, modulePath, serverResolved);
        }

        public override object InitializeLifetimeService() => null;
    }
}
