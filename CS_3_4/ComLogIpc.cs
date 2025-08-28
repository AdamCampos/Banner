using System;
using System.IO;
using System.IO.Pipes;
using System.Text;

namespace ComHookLib
{
    internal interface ILogSink : IDisposable
    {
        void WriteLine(string line);
    }

    internal sealed class ComLogIpc : ILogSink
    {
        private readonly string _pipeName;
        private readonly string _fallbackFile;
        private NamedPipeClientStream _pipe;
        private StreamWriter _writer;
        private readonly object _sync = new object();
        private volatile bool _useFile;

        public ComLogIpc(string pipeName, string fallbackFile)
        {
            _pipeName = string.IsNullOrWhiteSpace(pipeName) ? "FTAEPipe" : pipeName;
            _fallbackFile = string.IsNullOrWhiteSpace(fallbackFile)
                ? Path.Combine(Path.GetTempPath(), "ftaelog.log")
                : fallbackFile;

            TryConnectPipe();
        }

        private void TryConnectPipe()
        {
            try
            {
                _pipe = new NamedPipeClientStream(".", _pipeName, PipeDirection.Out, PipeOptions.Asynchronous);
                _pipe.Connect(50); // tenta rápido; se não houver servidor, cai para arquivo
                _writer = new StreamWriter(_pipe, new UTF8Encoding(false));
                _writer.AutoFlush = true;
                _useFile = false;
            }
            catch
            {
                // Fallback para arquivo
                string dir = Path.GetDirectoryName(_fallbackFile);
                if (string.IsNullOrEmpty(dir)) dir = Path.GetTempPath();
                Directory.CreateDirectory(dir);

                _writer = new StreamWriter(
                    new FileStream(_fallbackFile, FileMode.Append, FileAccess.Write, FileShare.ReadWrite),
                    new UTF8Encoding(false)
                );
                _writer.AutoFlush = true;
                _useFile = true;
            }
        }

        public void WriteLine(string line)
        {
            if (string.IsNullOrEmpty(line)) return;

            lock (_sync)
            {
                try
                {
                    _writer.WriteLine(line);
                }
                catch
                {
                    if (!_useFile)
                    {
                        // pipeline caiu; reabrir como arquivo
                        try { if (_writer != null) _writer.Dispose(); } catch { }
                        _writer = new StreamWriter(
                            new FileStream(_fallbackFile, FileMode.Append, FileAccess.Write, FileShare.ReadWrite),
                            new UTF8Encoding(false)
                        );
                        _writer.AutoFlush = true;
                        _useFile = true;
                        _writer.WriteLine(line);
                    }
                }
            }
        }

        public void Dispose()
        {
            try { if (_writer != null) _writer.Dispose(); } catch { }
            try { if (_pipe != null) _pipe.Dispose(); } catch { }
        }
    }
}
