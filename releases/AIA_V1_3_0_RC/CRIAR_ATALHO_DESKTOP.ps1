$ErrorActionPreference = "Stop"
$Base = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script = Join-Path $Base "AIA_V1_3_0_RC.py"
$Icon = Join-Path $Base "AIA_V1_3_0_RC.ico"

function Resolve-PythonExe {
    try {
        $p = (& py -3 -c "import sys; print(sys.executable)" 2>$null | Select-Object -First 1)
        if ($p -and (Test-Path $p)) { return $p.Trim() }
    } catch {}
    try {
        $p = (& python -c "import sys; print(sys.executable)" 2>$null | Select-Object -First 1)
        if ($p -and (Test-Path $p)) { return $p.Trim() }
    } catch {}
    return $null
}

$PythonExe = Resolve-PythonExe
if (-not $PythonExe) {
    Write-Host "[ERRO] Python 3 nao foi encontrado." -ForegroundColor Red
    Write-Host "Instale o Python 3 e execute INSTALAR_DEPENDENCIAS_AIA.bat antes de criar o atalho."
    exit 1
}
$PythonDir = Split-Path -Parent $PythonExe
$PythonW = Join-Path $PythonDir "pythonw.exe"
if (-not (Test-Path $PythonW)) { $PythonW = $PythonExe }

$Desktop = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $Desktop "AIA V1.3.0 RC.lnk"
$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $PythonW
$Shortcut.Arguments = '"' + $Script + '"'
$Shortcut.WorkingDirectory = $Base
$Shortcut.Description = "Autaza Inspection Analytics - AIA V1.3.0 RC"
if (Test-Path $Icon) { $Shortcut.IconLocation = $Icon + ",0" }
$Shortcut.Save()
Write-Host "Atalho criado com sucesso:" -ForegroundColor Green
Write-Host $ShortcutPath
