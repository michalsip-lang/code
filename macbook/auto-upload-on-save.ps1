param(
  [string]$RepoPath = "c:\Users\micha\Documents\GitHub\code",
  [string]$FileName = "hodnoceni_pracovnika",
  [string]$Branch = "main"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Set-Location $RepoPath

if (-not (Test-Path (Join-Path $RepoPath ".git"))) {
  throw "RepoPath není git repozitář: $RepoPath"
}

$targetPath = Join-Path $RepoPath $FileName
if (-not (Test-Path $targetPath)) {
  throw "Soubor neexistuje: $targetPath"
}

Write-Host "[AUTO-UPLOAD] Sleduji soubor: $targetPath"
Write-Host "[AUTO-UPLOAD] Branch: $Branch"

$fsw = New-Object System.IO.FileSystemWatcher
$fsw.Path = $RepoPath
$fsw.Filter = $FileName
$fsw.NotifyFilter = [IO.NotifyFilters]'LastWrite, Size, FileName'
$fsw.IncludeSubdirectories = $false
$fsw.EnableRaisingEvents = $true

$uploadAction = {
  try {
    Start-Sleep -Milliseconds 400

    Set-Location $using:RepoPath

    & git add -- $using:FileName

    $hasChanges = & git diff --cached --name-only
    if (-not $hasChanges) {
      return
    }

    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    & git commit -m "auto-save: $($using:FileName) $stamp"
    & git push origin $using:Branch

    Write-Host "[AUTO-UPLOAD] OK => $stamp"
  }
  catch {
    Write-Host "[AUTO-UPLOAD] CHYBA: $($_.Exception.Message)"
  }
}

$created = Register-ObjectEvent $fsw Created -Action $uploadAction
$changed = Register-ObjectEvent $fsw Changed -Action $uploadAction
$renamed = Register-ObjectEvent $fsw Renamed -Action $uploadAction

try {
  while ($true) {
    Start-Sleep -Seconds 1
  }
}
finally {
  Unregister-Event -SourceIdentifier $created.Name -ErrorAction SilentlyContinue
  Unregister-Event -SourceIdentifier $changed.Name -ErrorAction SilentlyContinue
  Unregister-Event -SourceIdentifier $renamed.Name -ErrorAction SilentlyContinue
  $fsw.Dispose()
}
