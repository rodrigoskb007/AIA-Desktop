$Desktop = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $Desktop "AIA V1.3.0 RC.lnk"
if (Test-Path $ShortcutPath) {
    Remove-Item $ShortcutPath -Force
    Write-Host "Atalho removido: $ShortcutPath" -ForegroundColor Green
} else {
    Write-Host "Nenhum atalho AIA V1.3.0 RC foi encontrado na area de trabalho."
}
