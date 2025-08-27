# Test-Injection.ps1 — DLL versionada por SHA, log fixo, verificação robusta; não fecha cliente no sucesso
# PowerShell 5.1 — Execute como Administrador

param(
    [string]$ProcessoAlvoPadrao = "DisplayClient",
    [string]$WindowTitleFilter,
    [int]$TargetProcId,
    [int]$Tentativas = 3,

    # Caminhos
    [string]$InjectorExe = "C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release\InjectorConsole.exe",
    [string]$DllFonte   = "C:\Projetos\VisualStudio\Banner\ComHookLib\bin\x86\Release\ComHookLib.dll",

    # Destinos legados (best-effort, podem falhar se arquivo estiver em uso)
    [string]$DllDestinoRelease = "C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release\ComHookLib.dll",
    [string]$DllDestinoExe     = "C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release\EXE\ComHookLib.dll",

    [string]$ArgsInjectorExtras = "",
    [switch]$OpenLog,          # abre log no final
    [switch]$AutoCloseOnFail   # fecha cliente automaticamente apenas se falhar
)

$ErrorActionPreference = 'Stop'
$Global:LogPath = "C:\Temp\ftaelog.log"
$Global:LastHashPath = "C:\Temp\ComHookLib.last.sha256"
$script:DllToInject = $null  # caminho EXATO (versionado por SHA) que será injetado

# =================== UTIL ===================

function Ensure-Dir {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

# Escrita COMPARTILHADA no log (FileShare.ReadWrite) + retentativa
function Write-LogLine {
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Text
    )

    if ($null -eq $Text) { $Text = "" }

    $dir = Split-Path -Parent $Global:LogPath
    Ensure-Dir $dir

    $enc = New-Object System.Text.UTF8Encoding($false)
    for ($i=0; $i -lt 5; $i++) {
        $fs = $null; $sw = $null
        try {
            $fs = New-Object System.IO.FileStream($Global:LogPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            $sw = New-Object System.IO.StreamWriter($fs, $enc)
            $sw.WriteLine($Text)
            $sw.Flush()
            break
        }
        catch {
            Start-Sleep -Milliseconds 150
            if ($i -eq 4) {
                Write-Warning ("Falha ao escrever no log '{0}': {1}" -f $Global:LogPath, $_.Exception.Message)
            }
        }
        finally {
            if ($sw) { $sw.Dispose() }
            if ($fs) { $fs.Dispose() }
        }
    }
}

# Reset de log sem lock exclusivo (Create + FileShare.ReadWrite)
function Reset-Log {
    $dir = Split-Path -Parent $Global:LogPath
    Ensure-Dir $dir

    $fs = $null
    try {
        $fs = New-Object System.IO.FileStream($Global:LogPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        # apenas truncar/criar
    } finally {
        if ($fs) { $fs.Dispose() }
    }
}

function W {
    param([string]$msg)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"), $msg
    Write-LogLine $line
    Write-Host $msg
}

function Test-PeArchitecture {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Arquivo não encontrado: $Path" }
    $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    try {
        $br = New-Object System.IO.BinaryReader($fs)
        $mz = $br.ReadBytes(2); if ($mz[0] -ne 0x4D -or $mz[1] -ne 0x5A) { throw "Assinatura MZ inválida em $Path" }
        $fs.Seek(0x3C,'Begin')|Out-Null; $peOffset = $br.ReadInt32()
        $fs.Seek($peOffset,'Begin')|Out-Null; $peSig = $br.ReadBytes(4)
        if ($peSig[0] -ne 0x50 -or $peSig[1] -ne 0x45) { throw "Assinatura PE inválida em $Path" }
        switch ($br.ReadUInt16()) { 0x014c {"x86"}; 0x8664 {"x64"}; default { "desconhecida" } }
    } finally { $fs.Dispose() }
}

function Get-FileSHA256 {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Arquivo não encontrado: $Path" }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $fs = [System.IO.File]::OpenRead($Path)
    try {
        ($sha.ComputeHash($fs) | ForEach-Object { $_.ToString("x2") }) -join ""
    } finally { $fs.Dispose(); $sha.Dispose() }
}

function Safe-Copy {
    param([string]$Src,[string]$Dst)
    try {
        $dir = Split-Path $Dst -Parent
        Ensure-Dir $dir
        Copy-Item -LiteralPath $Src -Destination $Dst -Force
        return $true
    } catch {
        W ("AVISO: não consegui copiar para '{0}': {1}" -f $Dst, $_.Exception.Message)
        return $false
    }
}

function Cleanup-OldHashedDlls {
    param([string]$Dir,[int]$Keep=5)
    if (-not (Test-Path -LiteralPath $Dir)) { return }
    $files = Get-ChildItem -LiteralPath $Dir -Filter "ComHookLib_*.dll" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    $toDel = $files | Select-Object -Skip $Keep
    foreach($f in $toDel){ try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue } catch {} }
}

function Ensure-X86-Dll-And-PrepareHashed {
    param([string]$Src,[string]$InjectorExePath,[string]$DllDestinoRelease,[string]$DllDestinoExe)

    if (-not (Test-Path -LiteralPath $Src)) { throw "DLL fonte não encontrada: $Src — faça o build x86 (Release) da ComHookLib." }
    $arch = Test-PeArchitecture $Src
    if ($arch -ne "x86") { throw "A DLL fonte não é x86 (detectado: $arch). Verifique o build." }

    $hash = Get-FileSHA256 $Src
    $prev = if (Test-Path -LiteralPath $Global:LastHashPath){ Get-Content -LiteralPath $Global:LastHashPath -ErrorAction SilentlyContinue } else { "" }
    if ($prev -and ($hash -eq $prev)) { W "AVISO: DLL fonte possui o mesmo SHA256 da última injeção (pode ser a mesma build)." }
    else { W "Nova DLL detectada (SHA256 diferente)." }
    Set-Content -LiteralPath $Global:LastHashPath -Value $hash -Encoding ASCII -Force

    # Caminho versionado por SHA dentro da pasta do injector
    $injDir = Split-Path -Path $InjectorExePath -Parent
    $hashedName = "ComHookLib_{0}.dll" -f $hash
    $hashedPath = Join-Path $injDir $hashedName

    # Copia para o arquivo versionado (não estará em uso)
    if (Safe-Copy -Src $Src -Dst $hashedPath) {
        W ("DLL x86 (versionada) copiada: {0}  ->  {1}" -f $Src, $hashedPath)
        $script:DllToInject = $hashedPath
    } else {
        throw "Falha ao copiar DLL para caminho versionado (hash)."
    }

    # Best-effort destinos legados (se travar, seguimos)
    [void](Safe-Copy -Src $Src -Dst $DllDestinoRelease)
    [void](Safe-Copy -Src $Src -Dst $DllDestinoExe)

    # Metadados
    try {
        $fvi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Src)
        W ("DLL Fonte: Version={0} | FileVersion={1} | ProductVersion={2} | LastWrite={3}" -f `
            $fvi.ProductVersion, $fvi.FileVersion, $fvi.ProductVersion, (Get-Item -LiteralPath $Src).LastWriteTime)
    } catch {
        W ("DLL Fonte: LastWrite={0} | SHA256={1}" -f (Get-Item -LiteralPath $Src).LastWriteTime, $hash)
    }

    Cleanup-OldHashedDlls -Dir $injDir -Keep 5
}

function Pick-ProcId {
    param([string]$ProcessoAlvoPadrao, [string]$WindowTitleFilter)
    $procs = Get-Process -Name $ProcessoAlvoPadrao -ErrorAction SilentlyContinue
    if ($WindowTitleFilter) { $procs = $procs | Where-Object { $_.MainWindowTitle -like "*$WindowTitleFilter*" } }
    if (-not $procs -or $procs.Count -eq 0) {
        W "Não encontrei '$ProcessoAlvoPadrao'$(if($WindowTitleFilter){" com título contendo '$WindowTitleFilter'"}). Listando processos..."
        Get-Process | Where-Object { -not $WindowTitleFilter -or ($_.MainWindowTitle -like "*$WindowTitleFilter*") } |
          Sort-Object ProcessName | Select-Object Id, ProcessName, MainWindowTitle | Out-Host
        return [int](Read-Host "Digite o PID de destino")
    }
    if ($procs.Count -eq 1) { return [int]$procs.Id }
    Write-Host "Foram encontrados vários '$ProcessoAlvoPadrao':"
    $i=0;$map=@{}; foreach ($p in $procs) { $i++; $map[$i]=$p.Id; "{0}. PID {1} - {2} {3}" -f $i,$p.Id,$p.ProcessName,($p.MainWindowTitle) | Out-Host }
    return [int]$map[[int](Read-Host "Escolha o índice")]
}

function Fechar-Clientes-FTView {
    $candidatos = @("DisplayClient","SEClient","ViewSEClient","FTViewSEClient","FTAEClient","FTClient","RSView32")
    foreach ($n in $candidatos) {
        Get-Process -Name $n -ErrorAction SilentlyContinue | ForEach-Object {
            W ("Encerrando '{0}' (PID={1})..." -f $_.ProcessName, $_.Id)
            $_ | Stop-Process -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-Injector {
    param([string]$InjectorExe,[int]$ProcId,[string]$LogPath,[string]$DllPath)

    $stdout = "C:\Temp\injector_stdout.log"
    $stderr = "C:\Temp\injector_stderr.log"
    Set-Content -LiteralPath $stdout -Value $null -Encoding UTF8 -Force
    Set-Content -LiteralPath $stderr -Value $null -Encoding UTF8 -Force

    # Assinatura: InjectorConsole <PID> <LogPath> [<DllPath>]
    $args = if ($DllPath) { @($ProcId, $LogPath, $DllPath) } else { @($ProcId, $LogPath) }

    $exit = 3
    try {
        $wd = Split-Path -Path $InjectorExe -Parent
        W ("Injector args (reais): {0}" -f ($args -join ' '))
        $p = Start-Process -FilePath $InjectorExe `
                           -ArgumentList $args `
                           -WorkingDirectory $wd `
                           -Wait -PassThru `
                           -WindowStyle Hidden `
                           -RedirectStandardOutput $stdout `
                           -RedirectStandardError  $stderr
        $exit = $p.ExitCode
        W ("Injector ExitCode={0}" -f $exit)
    } catch {
        W ("Exceção ao executar injector: {0}" -f $_.Exception.Message)
        return $false
    }

    # Registrar STDOUT/STDERR — qualquer erro aqui NÃO invalida a injeção
    try {
        Write-LogLine ""
        Write-LogLine "--- Injector STDOUT ---"
        Get-Content -LiteralPath $stdout -Encoding UTF8 | ForEach-Object { Write-LogLine $_ }
        Write-LogLine ""
        Write-LogLine "--- Injector STDERR ---"
        Get-Content -LiteralPath $stderr -Encoding UTF8 | ForEach-Object { Write-LogLine $_ }
    } catch {
        Write-Warning ("Falha ao registrar STDOUT/STDERR do Injector: {0}" -f $_.Exception.Message)
    }

    return ($exit -eq 0)
}

function Try-Inject {
    param([string]$InjectorExe,[int]$ProcId,[string]$DllPath)

    if (-not (Test-Path -LiteralPath $InjectorExe)) { throw "Injector não encontrado: $InjectorExe" }
    if (-not (Test-Path -LiteralPath $DllPath))    { throw "DLL para injeção não encontrada: $DllPath" }
    $injArch = Test-PeArchitecture $InjectorExe
    if ($injArch -ne "x86") { throw "Injector não é x86 (detectado: $injArch). Ajuste o build do InjectorConsole." }

    W "Injetando no PID $ProcId | Log: $Global:LogPath"
    if (Invoke-Injector -InjectorExe $InjectorExe -ProcId $ProcId -LogPath $Global:LogPath -DllPath $DllPath) {
        W ("Sucesso. PID={0} DLL={1}" -f $ProcId, $DllPath)
        return $true
    } else {
        W ("Falha de injeção. PID={0} DLL={1}" -f $ProcId, $DllPath)
        throw "Falha de injeção (todas as sintaxes testadas)."
    }
}

function Test-Process-Alive { param([int]$ProcId) try { Get-Process -Id $ProcId -ErrorAction Stop | Out-Null; $true } catch { $false } }

function Find-LoadedModule {
    param([int]$ProcId,[string]$ModuleName="ComHookLib.dll")
    # 1) via .NET
    try {
        $p = Get-Process -Id $ProcId -ErrorAction Stop
        foreach ($m in $p.Modules) { if ($m.ModuleName -ieq $ModuleName) { return $true } }
    } catch { }
    # 2) fallback via tasklist
    try {
        $out = & tasklist /FI "PID eq $ProcId" /M $ModuleName 2>$null
        if ($out -match [regex]::Escape($ModuleName)) { return $true }
    } catch { }
    return $false
}

function Confirm-RemoteHeartbeat {
    param([string]$Path = $Global:LogPath)
    try { (Get-Content -LiteralPath $Path -ErrorAction Stop -Tail 200) -match "\[REMOTE OK\]" } catch { $false }
}

# =================== MAIN ===================

Reset-Log

# 1) Prepara DLL x86 versionada por SHA
Ensure-X86-Dll-And-PrepareHashed -Src $DllFonte -InjectorExePath $InjectorExe -DllDestinoRelease $DllDestinoRelease -DllDestinoExe $DllDestinoExe
if (-not $script:DllToInject) { throw "Não foi possível determinar a DLL a injetar." }
W ("DLL a injetar: {0}" -f $script:DllToInject)

# 2) Resolver PID
if (-not $TargetProcId -or $TargetProcId -le 0) { $TargetProcId = Pick-ProcId -ProcessoAlvoPadrao $ProcessoAlvoPadrao -WindowTitleFilter $WindowTitleFilter }
if (-not ($TargetProcId -is [int]) -or $TargetProcId -le 0) { throw "PID inválido: '$TargetProcId'." }
W "Alvo: PID $TargetProcId"

# 3) Tentativas
for ($t = 1; $t -le $Tentativas; $t++) {
    W "Tentativa $t de $Tentativas..."
    try {
        $ok = Try-Inject -InjectorExe $InjectorExe -ProcId $TargetProcId -DllPath $script:DllToInject
        if ($ok) {
            Write-Host "✅ Injeção concluída com sucesso no PID $TargetProcId."

            if (Test-Process-Alive -ProcId $TargetProcId) {
                $hasDll = Find-LoadedModule -ProcId $TargetProcId -ModuleName "ComHookLib.dll"
                $hb = Confirm-RemoteHeartbeat

                if ($hasDll) {
                    W "Verificação: ComHookLib.dll ESTÁ carregada no PID $TargetProcId."
                } elseif ($hb) {
                    W "Verificação: [REMOTE OK] detectado no log."
                } else {
                    W "Verificação: NÃO encontrei ComHookLib.dll no PID e nenhum [REMOTE OK] no log."
                }
            } else {
                W "Processo $TargetProcId encerrou antes da verificação."
            }

            if ($OpenLog) { notepad $Global:LogPath | Out-Null }
            return
        }
    } catch {
        W ("Falha na injeção: {0}" -f $_.Exception.Message)
        if ($AutoCloseOnFail) {
            Fechar-Clientes-FTView
        } else {
            $resp = Read-Host "Deseja fechar o cliente antes de tentar novamente? (S/N)"
            if ($resp -match '^(s|S|y|Y)') { Fechar-Clientes-FTView }
        }
        Start-Sleep -Seconds 2
        $TargetProcId = Pick-ProcId -ProcessoAlvoPadrao $ProcessoAlvoPadrao -WindowTitleFilter $WindowTitleFilter
        if (-not ($TargetProcId -is [int]) -or $TargetProcId -le 0) { throw "PID inválido: '$TargetProcId'." }
    }
}

Write-Error "❌ Não foi possível injetar após $Tentativas tentativas."
if ($OpenLog) { notepad $Global:LogPath | Out-Null }
exit 1
