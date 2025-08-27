using System;

namespace ComHookLib
{
    /// <summary>
    /// Etapa 1/4: placeholder de inicialização para hooks COM.
    /// Nesta etapa ainda NÃO instalamos hooks de CoCreateInstance/Ex.
    /// </summary>
    internal static class ComHooks
    {
        private static bool _initialized;
        private static ILogger _logger;

        public static void TryInitialize(ILogger logger)
        {
            if (_initialized) return;
            _logger = logger;
            if (_logger != null)
                _logger.Info("ComHooks (etapa 1/4): inicializado em modo placeholder (sem hooks COM).");
            _initialized = true;
        }
    }
}
