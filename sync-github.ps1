param(
  [string]$RepoPath = $PSScriptRoot,
  [string]$Branch = "main"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Invoke-Git {
  param([string[]]$Arguments)

  & git @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "Příkaz git selhal: git $($Arguments -join ' ')"
  }
}

Set-Location $RepoPath

if (-not (Test-Path (Join-Path $RepoPath ".git"))) {
  throw "Cesta není Git repozitář: $RepoPath"
}

# Neprovádí synchronizaci přes nevyřešený konflikt.
$conflicts = & git diff --name-only --diff-filter=U
if ($conflicts) {
  throw "Synchronizace zastavena: nevyřešené konflikty v repozitáři."
}

Invoke-Git @("add", "--all")
$changes = & git diff --cached --quiet
if ($LASTEXITCODE -eq 1) {
  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Invoke-Git @("commit", "-m", "Automatická synchronizace: $timestamp")
}
elseif ($LASTEXITCODE -ne 0) {
  throw "Nelze ověřit připravené změny."
}

Invoke-Git @("fetch", "origin", $Branch)

$behind = & git rev-list --count "HEAD..origin/$Branch"
if ([int]$behind -gt 0) {
  Invoke-Git @("pull", "--rebase", "origin", $Branch)
}

$ahead = & git rev-list --count "origin/$Branch..HEAD"
if ([int]$ahead -gt 0) {
  Invoke-Git @("push", "origin", $Branch)
}

Write-Output "GitHub synchronizace dokončena: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"