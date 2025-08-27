# Test-Injection.ps1  (PS 5.1)
# - Modo Attach: injeta numa PID já em execução
# - Modo Smoke : cria um alvo 32-bit (wscript.exe + VBS) pra validar hooks e eventos COM
# Observações:
#   * Sem uso de $pid
#   * Sem Stop-Job -Force (incompatível no PS 5.1)
#   * Inclui esperas configuráveis para permitir interação com o banner (ciclos ~2 min)

param(
    [ValidateSet('Attach','Smoke')]
    [string]$Mode = 'Attach',

    # Usado no modo Attach
    [int]$TargetPid,

    # Caminhos (ajuste se necessário)
    [string]$InjectorExe = 'C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release\InjectorConsole.exe',
    [string]$DllSource   = 'C:\Projetos\VisualStudio\Banner\ComHookLib\bin\x86\Release\ComHookLib.dll',
    [string]$BinOutDir   = 'C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release',
    [string]$LogPath     = 'C:\Temp\ftaelog.log',

    # Esperas
    [int]$TailSeconds       = 10,    # "espiadinha" inicial do log logo após injetar
    [int]$ComTimeoutSeconds = 180,   # quanto esperar por eventos COM (aumente p/ ciclo > 2min)
    [int]$StayAttachedSeconds = 0    # opcional: tail prolongado p/ você interagir no banner (0 = desliga)
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Get-FileSha256 {
    param([string]$Path)
    if (!(Test-Path $Path)) { return $null }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $fs  = [System.IO.File]::OpenRead($Path)
    try { return ($sha.ComputeHash($fs) | ForEach-Object { $_.ToString('x2') }) -join '' }
    finally { $fs.Dispose(); $sha.Dispose() }
}

function New-VersionedCopy {
    param([string]$Source, [string]$OutDir)
    if (!(Test-Path $Source)) { throw "DLL fonte não encontrada: $Source" }
    if (!(Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }

    $hash = Get-FileSha256 -Path $Source
    if ($hash) { Write-Warning "DLL fonte possui o mesmo SHA256 da última injeção (pode ser a mesma build)." }

    $target = Join-Path $OutDir ("ComHookLib_{0}.dll" -f $hash)
    Copy-Item -LiteralPath $Source -Destination $target -Force

    $fi = Get-Item $Source
    Write-Host ("DLL x86 (versionada) copiada: {0}  ->  {1}" -f $Source, $target)
    Write-Host ("DLL Fonte: Version={0} | FileVersion={1} | ProductVersion={2} | LastWrite={3}" -f `
        ($fi.VersionInfo.ProductVersion), ($fi.VersionInfo.FileVersion), ($fi.VersionInfo.ProductVersion), $fi.LastWriteTime)
    return $target
}

function Start-LogsTail {
    param([string]$Path,[int]$Seconds)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $Seconds) {
        if (Test-Path $Path) {
            # -Raw não é necessário aqui; preferimos últimas linhas
            Get-Content -LiteralPath $Path -Tail 15
        }
        Start-Sleep -Milliseconds 500
    }
    $sw.Stop()
}

function Ensure-RemoteOk {
    param([string]$Path)
    if (!(Test-Path $Path)) { throw "Log não encontrado: $Path" }
    $content = Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue
    $ok1 = $content -match '\[REMOTE OK\] Hooks \(UI \+ COM\) instalados\.'
    $ok2 = $content -match 'COM hooks instalados com sucesso \(CoCreateInstance/Ex, CoGetClassObject\)\.'
    $ok3 = $content -match '\[REMOTE OK\] RemoteEntry iniciado\.'
    if (!($ok1 -or $ok2 -or $ok3)) {
        throw "Confirmação de instalação não encontrada ainda em $Path"
    }
    Write-Host "Verificação: instalação dos hooks confirmada."
}

function Wait-ForComActivity {
    param([string]$Path,[int]$TimeoutSeconds=180)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (Test-Path $Path) {
            $content = Get-Content -LiteralPath $Path -Raw -ErrorAction SilentlyContinue
            if ($content -match '"api":"Co(CreateInstance|CreateInstanceEx|GetClassObject)"') {
                Write-Host "✔ Eventos COM detectados."
                return
            }
        }
        Start-Sleep -Milliseconds 300
    }
    Write-Warning "Nenhum evento COM visto em ${TimeoutSeconds}s. Gere interação no banner ou aumente o timeout."
}

function Tail-ForInteraction {
    param([string]$Path,[int]$Seconds)
    if ($Seconds -le 0) { return }
    Write-Host ("-- Mantendo tail por {0}s para permitir interação no banner..." -f $Seconds)
    Start-LogsTail -Path $Path -Seconds $Seconds
}

function Invoke-Injection {
    param([int]$TargetPid, [string]$DllToInject, [string]$LogPath)
    Write-Host "Alvo: PID $TargetPid"
    Write-Host "Tentativa 1 de 1..."
    Write-Host "Injetando no PID $TargetPid | Log: $LogPath"
    Write-Host "Injector args (reais): $TargetPid $LogPath $DllToInject"
    & $InjectorExe $TargetPid $LogPath $DllToInject | Write-Output
    if ($LASTEXITCODE -ne 0) { throw "Injector retornou código $LASTEXITCODE" }
    Write-Host "Injector ExitCode=$LASTEXITCODE"
}

# --- Preparar destino de log (apagar, truncar ou redirecionar) ---
if (Test-Path $LogPath) {
    try {
        Remove-Item $LogPath -Force -ErrorAction Stop
    } catch {
        Write-Warning "Log em uso, farei TRUNCATE ao invés de deletar: $LogPath"
        try {
            $fs=[System.IO.File]::Open($LogPath,[System.IO.FileMode]::OpenOrCreate,[System.IO.FileAccess]::ReadWrite,[System.IO.FileShare]::ReadWrite)
            $fs.SetLength(0); $fs.Flush(); $fs.Dispose()
        } catch {
            $new = [IO.Path]::Combine([IO.Path]::GetDirectoryName($LogPath),("ftaelog_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")))
            Write-Warning "Não consegui truncar. Redirecionando log para: $new"
            $script:LogPath = $new
        }
    }
}

# --- Fluxo principal ---
$DllToInject = New-VersionedCopy -Source $DllSource -OutDir $BinOutDir

switch ($Mode) {

'Attach' {
    if (-not $TargetPid) { throw "Modo Attach requer -TargetPid <processo alvo>." }
    Invoke-Injection -TargetPid $TargetPid -DllToInject $DllToInject -LogPath $LogPath

    # 1) Espiar log imediatamente após a injeção
    Start-LogsTail -Path $LogPath -Seconds $TailSeconds

    # 2) Confirmar hooks remotos
    Ensure-RemoteOk -Path $LogPath

    # 3) Aguardar atividade COM (útil quando o ciclo do banner demora)
    Wait-ForComActivity -Path $LogPath -TimeoutSeconds $ComTimeoutSeconds

    # 4) Opcional: manter tail para você interagir livremente
    Tail-ForInteraction -Path $LogPath -Seconds $StayAttachedSeconds

    Write-Host "✅ Injeção (Attach) concluída."
}

'Smoke' {
    # Gera alvo 32-bit para casar com DLL x86
    $WScript32 = "$env:WINDIR\SysWOW64\wscript.exe"
    if (!(Test-Path $WScript32)) { throw "wscript.exe (x86) não encontrado em $WScript32" }

    # VBS de estímulo COM (demora um pouco e cria vários objetos COM)
    $VbsPath = Join-Path $env:TEMP 'com_smoke_test.vbs'
@'
WScript.Echo "VBS iniciado, aguardando 4000 ms para injeção..."
WScript.Sleep 4000
Set d = CreateObject("Scripting.Dictionary")
d.Add "k1", "v1"
Set sh = CreateObject("WScript.Shell")
sh.Environment("PROCESS")("ABC") = "123"
For i = 1 To 3
  Set d2 = CreateObject("Scripting.Dictionary")
  Set sh2 = CreateObject("WScript.Shell")
Next
WScript.Echo "Teste COM concluído."
WScript.Sleep 2000
'@ | Set-Content -Path $VbsPath -Encoding ASCII

    $proc = Start-Process -FilePath $WScript32 -ArgumentList "`"$VbsPath`"" -PassThru

    Invoke-Injection -TargetPid $proc.Id -DllToInject $DllToInject -LogPath $LogPath
    Start-LogsTail -Path $LogPath -Seconds $TailSeconds
    Ensure-RemoteOk -Path $LogPath
    Wait-ForComActivity -Path $LogPath -TimeoutSeconds $ComTimeoutSeconds

    if (-not $proc.HasExited) { $proc | Stop-Process -Force }
    Write-Host "✅ Injeção (Smoke) concluída."
}
}

<#
Exemplos de uso:

$ts = Get-Date -Format "yyyyMMdd_HHmmss"

# Attach ao seu processo alvo + 2min para aguardar eventos COM + 2min extras de tail para você interagir
.\Test-Injection.ps1 -Mode Attach -TargetPid 23880 `
  -InjectorExe 'C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release\InjectorConsole.exe' `
  -DllSource  'C:\Projetos\VisualStudio\Banner\ComHookLib\bin\x86\Release\ComHookLib.dll' `
  -BinOutDir  'C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release' `
  -LogPath    "C:\Temp\ftaelog_$ts.log" `
  -ComTimeoutSeconds 150 `
  -StayAttachedSeconds 120

# Smoke test rápido
.\Test-Injection.ps1 -Mode Smoke `
  -InjectorExe 'C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release\InjectorConsole.exe' `
  -DllSource  'C:\Projetos\VisualStudio\Banner\ComHookLib\bin\x86\Release\ComHookLib.dll' `
  -BinOutDir  'C:\Projetos\VisualStudio\Banner\InjectorConsole\bin\x86\Release' `
  -LogPath    "C:\Temp\ftaelog_smoke_$ts.log" `
  -ComTimeoutSeconds 20
#>
