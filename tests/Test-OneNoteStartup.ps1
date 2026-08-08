param(
    [Parameter(Mandatory = $true)]
    [string]$SetupPath,
    [ValidateSet("x86", "x64")]
    [string]$Platform = "x64",
    [string]$OneNotePath,
    [ValidateRange(15, 120)]
    [int]$ObservationSeconds = 30
)

$ErrorActionPreference = "Stop"
$setup = (Resolve-Path $SetupPath).Path
$addInAppId = "CCF3B64A-9580-4D44-BDCE-49BFE95AC510"

function Get-AddInSurrogate {
    @(Get-CimInstance Win32_Process -Filter "Name='dllhost.exe'" |
        Where-Object { $_.CommandLine -match $addInAppId })
}

if (-not $OneNotePath) {
    $programFiles = if ($Platform -eq "x86") {
        ${env:ProgramFiles(x86)}
    } else {
        $env:ProgramFiles
    }
    $OneNotePath = Join-Path $programFiles "Microsoft Office\root\Office16\ONENOTE.EXE"
}
if (-not (Test-Path $OneNotePath)) {
    throw "OneNote Desktop wurde nicht gefunden: $OneNotePath"
}
if (Get-Process ONENOTE -ErrorAction SilentlyContinue) {
    throw "OneNote muss vor dem Host-Regressionstest vollstandig beendet sein."
}
if ((Get-AddInSurrogate).Count -ne 0) {
    throw "Ein alter Schnellbausteine-COM-Surrogatprozess lauft noch."
}

$install = Start-Process $setup -ArgumentList @(
    "/VERYSILENT",
    "/SUPPRESSMSGBOXES",
    "/NORESTART",
    "/ONENOTEBITNESS=$Platform"
) -Wait -PassThru
if ($install.ExitCode -ne 0) {
    throw "Setup wurde mit Exitcode $($install.ExitCode) beendet."
}

$started = Get-Date
$oneNote = Start-Process $OneNotePath -PassThru
$ready = $false
$respondingSamples = 0

try {
    for ($second = 1; $second -le $ObservationSeconds; $second++) {
        Start-Sleep -Seconds 1
        $oneNote.Refresh()
        if ($oneNote.HasExited) {
            throw "OneNote wurde vorzeitig mit Exitcode $($oneNote.ExitCode) beendet."
        }
        if ($oneNote.Responding) {
            $respondingSamples++
        }
        if ($oneNote.Responding -and $oneNote.MainWindowHandle -ne 0) {
            $ready = $true
        }
    }

    $registryView = if ($Platform -eq "x86") {
        [Microsoft.Win32.RegistryView]::Registry32
    } else {
        [Microsoft.Win32.RegistryView]::Registry64
    }
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser, $registryView)
    try {
        $addIn = $base.OpenSubKey(
            "Software\Microsoft\Office\OneNote\AddIns\Schnellbausteine.OneNoteAddIn")
        if (-not $addIn) { throw "OneNote-Add-in-Registrierung fehlt." }
        if ($addIn.GetValue("LoadBehavior") -ne 3) {
            throw "OneNote hat LoadBehavior nach dem Start verandert."
        }
        $addIn.Close()
    }
    finally {
        $base.Close()
    }

    $surrogate = @(Get-AddInSurrogate | Where-Object { $_.CreationDate -ge $started })
    if ($surrogate.Count -ne 1) {
        throw "Der erwartete Schnellbausteine-COM-Surrogatprozess fehlt."
    }

    $hostErrors = @(Get-WinEvent -FilterHashtable @{
        LogName = "Application"
        StartTime = $started
    } -ErrorAction SilentlyContinue | Where-Object {
        $_.ProviderName -in @(".NET Runtime", "Application Error", "Application Hang") -and
        $_.Message -match "ONENOTE|DllHost"
    })
    if ($hostErrors.Count -ne 0) {
        throw "Seit dem OneNote-Start wurden $($hostErrors.Count) Hostfehler protokolliert."
    }
    if (-not $ready -or -not $oneNote.Responding) {
        throw "OneNote hat kein dauerhaft responsives Hauptfenster erreicht."
    }
    if ($respondingSamples -ne $ObservationSeconds) {
        throw "OneNote war nicht wahrend des gesamten Beobachtungszeitraums responsiv."
    }

    Write-Host "Realer OneNote-Start mit Schnellbausteine war erfolgreich."
}
finally {
    if (-not $oneNote.HasExited) {
        $oneNote.CloseMainWindow() | Out-Null
        if (-not $oneNote.WaitForExit(15000)) {
            Write-Warning (
                "OneNote lauft nach dem Schliessen des Fensters im Infobereich weiter. " +
                "Beenden Sie es dort vollstandig.")
        }
    }
}
