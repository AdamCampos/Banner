// ComHookLib - MultiSink.cs
// Encadeia 2+ sinks existentes.

using System;
using System.Collections.Generic;

namespace ComHookLib
{
    internal sealed class MultiSink : ILogSink, IDisposable
    {
        private readonly List<ILogSink> _sinks;

        internal MultiSink(params ILogSink[] sinks)
        {
            _sinks = new List<ILogSink>(sinks ?? new ILogSink[0]);
        }

        public void WriteLine(string line)
        {
            for (int i = 0; i < _sinks.Count; i++)
                _sinks[i].WriteLine(line);
        }

        public void Dispose()
        {
            for (int i = 0; i < _sinks.Count; i++)
            {
                var d = _sinks[i] as IDisposable;
                if (d != null) d.Dispose();
            }
        }
    }
}
