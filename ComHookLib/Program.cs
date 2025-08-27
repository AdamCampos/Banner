using System;
using System.Diagnostics;
using System.IO;
using EasyHook;

class Injector
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Uso: InjectorConsole <PID> <LogPath>");
            Console.WriteLine(@"Ex.: InjectorConsole 17656 C:\Temp\com_activations_remote.log");
            return;
        }

        int pid = int.Parse(args[0]);
        string logPath = args[1];

        var dir = Path.GetDirectoryName(logPath);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        // NÃO referenciar ComHookLib aqui; sem IPC:
        string channel = null;

        // Caminhos das DLLs a injetar (as duas arquiteturas):
        string dll32 = Path.GetFullPath("ComHookLib32.dll");
        string dll64 = Path.GetFullPath("ComHookLib64.dll");

        if (!File.Exists(dll32)) { Console.WriteLine("ERRO: ComHookLib32.dll não encontrado: " + dll32); return; }
        if (!File.Exists(dll64)) { Console.WriteLine("ERRO: ComHookLib64.dll não encontrado: " + dll64); return; }

        Console.WriteLine($"Injetando no PID {pid} com log em: {logPath}");
        try
        {
            RemoteHooking.Inject(
                pid,
                InjectionOptions.Default,
                dll32,  // x86
                dll64,  // x64
                channel,
                logPath
            );
            Console.WriteLine("Injetado. Pressione Enter para encerrar...");
            Console.ReadLine();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Falha na injeção: " + ex);
        }
    }
}
