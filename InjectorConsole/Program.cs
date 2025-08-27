using System;
using System.IO;
using EasyHook;
using ComHookLib; // <- ESTE é o namespace correto que contém ComLogIpc

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
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        // ⚠️ Forçar o tipo com namespace totalmente qualificado evita erros de resolução
        string channel = null;
        RemoteHooking.IpcCreateServer<ComHookLib.ComLogIpc>(
            ref channel,
            System.Runtime.Remoting.WellKnownObjectMode.Singleton);

        // Caminho da DLL a injetar
        string hookLib = Path.GetFullPath("ComHookLib.dll");
        if (!File.Exists(hookLib))
        {
            Console.WriteLine("ERRO: ComHookLib.dll não encontrado: " + hookLib);
            return;
        }

        Console.WriteLine($"Injetando no PID {pid} com log em: {logPath}");
        try
        {
            RemoteHooking.Inject(
                pid,
                InjectionOptions.Default,
                hookLib, hookLib,    // x86/x64: use a DLL correta para a arquitetura do alvo
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
