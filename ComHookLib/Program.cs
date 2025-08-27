using System;
using System.IO;
using EasyHook;
using ComHookLib; // referência ao projeto ComHookLib para o tipo ComLogIpc

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

        // Garante a pasta do log passado na linha de comando (apenas informativo)
        var dir = Path.GetDirectoryName(logPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        // Cria o servidor IPC e obtém o 'channel' que será passado ao remoto
        string channel = null;
        RemoteHooking.IpcCreateServer<ComLogIpc>(
            ref channel,
            System.Runtime.Remoting.WellKnownObjectMode.Singleton);

        // Caminhos das DLLs a injetar (32 e 64 bits)
        string dll32 = Path.GetFullPath("ComHookLib32.dll");
        string dll64 = Path.GetFullPath("ComHookLib64.dll");

        if (!File.Exists(dll32))
        {
            Console.WriteLine("ERRO: ComHookLib32.dll não encontrado: " + dll32);
            return;
        }
        if (!File.Exists(dll64))
        {
            Console.WriteLine("ERRO: ComHookLib64.dll não encontrado: " + dll64);
            return;
        }

        Console.WriteLine($"Injetando no PID {pid} (canal IPC: {channel})");
        try
        {
            // *** PASSA SOMENTE 1 ARGUMENTO EXTRA (channel) ***
            RemoteHooking.Inject(
                pid,
                InjectionOptions.Default,
                dll32,  // x86
                dll64,  // x64
                channel // <-- casa com RemoteEntry.Run(context, string)
            );

            Console.WriteLine("Injetado. Pressione Enter para encerrar o injector...");
            Console.ReadLine();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Falha na injeção: " + ex);
        }
    }
}
