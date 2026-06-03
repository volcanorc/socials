$ErrorActionPreference = 'Stop'

Set-Location -LiteralPath $PSScriptRoot

$runtimeCache = Join-Path $PSScriptRoot '.node-cache'
$portableVersion = 'v22.16.0'
$portableFolder = "node-$portableVersion-win-x64"
$portableRoot = Join-Path $runtimeCache $portableFolder
$portableExe = Join-Path $portableRoot 'node.exe'

function Get-SystemNode {
  $candidates = @()

  $commandNode = Get-Command node -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandType -eq 'Application' -and $_.Source } |
    Select-Object -ExpandProperty Source -First 1
  if ($commandNode) {
    $candidates += $commandNode
  }

  if ($env:ProgramFiles) {
    $candidates += (Join-Path $env:ProgramFiles 'nodejs\node.exe')
  }

  if (${env:ProgramFiles(x86)}) {
    $candidates += (Join-Path ${env:ProgramFiles(x86)} 'nodejs\node.exe')
  }

  $candidates += @(
    'C:\Program Files\nodejs\node.exe',
    'C:\Program Files (x86)\nodejs\node.exe'
  )

  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path -LiteralPath $candidate)) {
      return $candidate
    }
  }

  return $null
}

function Ensure-PortableNode {
  if (Test-Path -LiteralPath $portableExe) {
    return $portableExe
  }

  New-Item -ItemType Directory -Force -Path $runtimeCache | Out-Null

  $zipPath = Join-Path $runtimeCache "$portableFolder.zip"
  $downloadUrl = "https://nodejs.org/dist/$portableVersion/$portableFolder.zip"

  Write-Host "Downloading portable Node.js $portableVersion..."
  try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing
  } catch {
    throw "Unable to download portable Node.js from $downloadUrl. Install Node.js manually or try again with internet access."
  }

  try {
    if (Test-Path -LiteralPath $portableRoot) {
      Remove-Item -LiteralPath $portableRoot -Recurse -Force
    }
    Expand-Archive -Path $zipPath -DestinationPath $runtimeCache -Force
  } catch {
    throw "Portable Node.js downloaded but could not be unpacked. Delete '.node-cache' and try again."
  }

  if (Test-Path -LiteralPath $portableExe) {
    return $portableExe
  }

  $found = Get-ChildItem -Path $runtimeCache -Recurse -Filter node.exe |
    Where-Object { $_.FullName -like "*$portableFolder*" } |
    Select-Object -First 1

  if ($found) {
    return $found.FullName
  }

  throw "Portable Node.js setup failed because node.exe was not found after extraction."
}

function Get-NodeExecutable {
  if (Test-Path -LiteralPath $portableExe) {
    return $portableExe
  }

  $systemNode = Get-SystemNode
  if ($systemNode) {
    return $systemNode
  }

  return Ensure-PortableNode
}

try {
  $nodeExe = Get-NodeExecutable
  Write-Host "Using Node runtime: $nodeExe"
  & $nodeExe (Join-Path $PSScriptRoot 'server.js')
} catch {
  Write-Host ""
  Write-Host "Launcher failed:"
  Write-Host $_.Exception.Message
  Write-Host ""
  Read-Host "Press Enter to close"
}
