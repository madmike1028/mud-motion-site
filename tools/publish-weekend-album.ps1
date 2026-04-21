[CmdletBinding()]
param(
  [string]$SourceFolder,
  [string]$EventDate,
  [string]$AlbumTitle,
  [string]$DestinationRoot = "",
  [int]$MaxLongEdge = 2200,
  [ValidateRange(1, 100)]
  [int]$JpegQuality = 82,
  [switch]$Force,
  [switch]$Quiet,
  [switch]$SkipGit,
  [switch]$SkipOpenFolder
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $PSCommandPath }
if (-not $DestinationRoot) {
  $DestinationRoot = Join-Path $scriptRoot "..\media\photo-albums"
}

$windowTitle = "Weekend Album Publisher"
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptRoot ".."))
$destinationRootFull = [System.IO.Path]::GetFullPath($DestinationRoot)
$exportScript = Join-Path $scriptRoot "export-photo-albums.ps1"
$metadataFileName = ".mud-motion-weekend.json"

function Show-Info {
  param([string]$Message)

  if ($Quiet) {
    Write-Output $Message
    return
  }

  [System.Windows.Forms.MessageBox]::Show(
    $Message,
    $windowTitle,
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
  ) | Out-Null
}

function Show-ErrorBox {
  param([string]$Message)

  if ($Quiet) {
    Write-Error $Message
    return
  }

  [System.Windows.Forms.MessageBox]::Show(
    $Message,
    $windowTitle,
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error
  ) | Out-Null
}

function Select-Folder {
  param(
    [string]$Description,
    [string]$SelectedPath
  )

  $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
  $dialog.Description = $Description
  $dialog.ShowNewFolderButton = $false

  if ($SelectedPath -and (Test-Path -LiteralPath $SelectedPath)) {
    $dialog.SelectedPath = [System.IO.Path]::GetFullPath($SelectedPath)
  }

  try {
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
      return $null
    }

    return $dialog.SelectedPath
  } finally {
    $dialog.Dispose()
  }
}

function Normalize-AlbumName {
  param([string]$Name)

  if (-not $Name) {
    return ""
  }

  $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
  $builder = New-Object System.Text.StringBuilder

  foreach ($char in $Name.ToCharArray()) {
    if ($invalidChars -contains $char) {
      [void]$builder.Append(" ")
    } else {
      [void]$builder.Append($char)
    }
  }

  return (($builder.ToString() -replace "\s+", " ").Trim().TrimEnd("."))
}

function Get-GitCommand {
  $gitCommand = Get-Command git.exe -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($gitCommand) {
    return $gitCommand.Source
  }

  $desktopRoot = Join-Path $env:LOCALAPPDATA "GitHubDesktop"
  if (Test-Path -LiteralPath $desktopRoot) {
    $candidate = Get-ChildItem -LiteralPath $desktopRoot -Directory -Filter "app-*" |
      Sort-Object LastWriteTime -Descending |
      ForEach-Object { Join-Path $_.FullName "resources\app\git\cmd\git.exe" } |
      Where-Object { Test-Path -LiteralPath $_ } |
      Select-Object -First 1

    if ($candidate) {
      return $candidate
    }
  }

  throw "Git was not found. Please install GitHub Desktop on this computer first."
}

function Invoke-Git {
  param(
    [string]$GitExe,
    [string[]]$Arguments,
    [switch]$AllowFailure
  )

  $output = @(& $GitExe -C $repoRoot @Arguments 2>&1)
  $exitCode = $LASTEXITCODE
  if ($exitCode -ne 0 -and (-not $AllowFailure)) {
    $message = ($output | ForEach-Object { $_.ToString() }) -join [Environment]::NewLine
    if (-not $message) {
      $message = "Git failed with exit code $exitCode."
    }
    throw $message
  }

  return [pscustomobject]@{
    ExitCode = $exitCode
    Output = $output | ForEach-Object { $_.ToString() }
  }
}

function Get-RelativeRepoPath {
  param([string]$Path)

  $baseFullPath = [System.IO.Path]::GetFullPath($repoRoot).TrimEnd("\")
  $targetFullPath = [System.IO.Path]::GetFullPath($Path)
  $baseUri = New-Object System.Uri($baseFullPath + "\")
  $targetUri = New-Object System.Uri($targetFullPath)
  $relativeUri = $baseUri.MakeRelativeUri($targetUri)
  return [System.Uri]::UnescapeDataString($relativeUri.ToString()).Replace("\", "/")
}

function Get-StatusPath {
  param([string]$StatusLine)

  $text = ($StatusLine -replace '^[ MARCUD?!]{2}\s+', '').Trim()
  if ($text -match ' -> ') {
    return ($text -split ' -> ')[-1].Trim()
  }

  return $text
}

function Read-Metadata {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) {
    return $null
  }

  return Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
}

function Write-Metadata {
  param(
    [string]$Path,
    [object]$Data
  )

  $Data | ConvertTo-Json | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Get-ValidatedEventDate {
  param([string]$Value)

  try {
    $parsed = [System.DateTime]::ParseExact($Value, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
    return $parsed.ToString("yyyy-MM-dd")
  } catch {
  }

  throw "Use the date like this: 2026-04-11"
}

try {
  if (-not (Test-Path -LiteralPath $exportScript)) {
    throw "Could not find the export tool at $exportScript"
  }

  if (-not (Test-Path -LiteralPath $destinationRootFull)) {
    New-Item -ItemType Directory -Path $destinationRootFull -Force | Out-Null
  }

  if (-not $SourceFolder) {
    $SourceFolder = Select-Folder -Description "Pick the race weekend folder. It should have Saturday and Sunday inside it." -SelectedPath $env:USERPROFILE
    if (-not $SourceFolder) {
      return
    }
  }

  $sourceFolderFull = [System.IO.Path]::GetFullPath($SourceFolder)
  if (-not (Test-Path -LiteralPath $sourceFolderFull)) {
    throw "Source folder not found: $sourceFolderFull"
  }

  if ($sourceFolderFull.StartsWith($repoRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Please pick the original weekend folder outside the website folder."
  }

  $metadataPath = Join-Path $sourceFolderFull $metadataFileName
  $metadata = Read-Metadata -Path $metadataPath

  if (-not $EventDate) {
    if ($metadata -and $metadata.eventDate) {
      $EventDate = $metadata.eventDate
    } else {
      $EventDate = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Type the race weekend date like 2026-04-11. This date controls the homepage order.",
        $windowTitle,
        (Get-Date).ToString("yyyy-MM-dd")
      )

      if (-not $EventDate) {
        return
      }
    }
  }

  $EventDate = Get-ValidatedEventDate -Value $EventDate

  if (-not $AlbumTitle) {
    if ($metadata -and $metadata.albumTitle) {
      $AlbumTitle = $metadata.albumTitle
    } else {
      $AlbumTitle = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Type the race weekend name exactly how it should show on the website.",
        $windowTitle,
        (Split-Path -Leaf $sourceFolderFull)
      )

      if (-not $AlbumTitle) {
        return
      }
    }
  }

  $cleanAlbumTitle = Normalize-AlbumName -Name $AlbumTitle
  if (-not $cleanAlbumTitle) {
    throw "Album name is blank after cleaning it up. Please try again."
  }

  $albumFolderName = "$EventDate $cleanAlbumTitle"
  if ($metadata -and $metadata.albumFolderName -and (-not $PSBoundParameters.ContainsKey("EventDate")) -and (-not $PSBoundParameters.ContainsKey("AlbumTitle"))) {
    $albumFolderName = $metadata.albumFolderName
  }

  $destinationFolder = Join-Path $destinationRootFull $albumFolderName
  $relativeAlbumPath = Get-RelativeRepoPath -Path $destinationFolder

  Write-Metadata -Path $metadataPath -Data ([pscustomobject]@{
    eventDate = $EventDate
    albumTitle = $cleanAlbumTitle
    albumFolderName = $albumFolderName
    updatedAt = (Get-Date).ToString("s")
  })

  $exportParams = @{
    SourceRoot = $sourceFolderFull
    DestinationRoot = $destinationFolder
    MaxLongEdge = $MaxLongEdge
    JpegQuality = $JpegQuality
  }

  if ($Force) {
    $exportParams.Force = $true
  }

  $output = @(& $exportScript @exportParams 2>&1)

  if (-not $SkipOpenFolder) {
    Start-Process explorer.exe $destinationFolder
  }

  if ($SkipGit) {
    Show-Info @"
Weekend album is ready.

Saved here:
$destinationFolder

Git push was skipped because SkipGit was turned on.
"@
    return
  }

  $gitExe = Get-GitCommand
  $statusAll = Invoke-Git -GitExe $gitExe -Arguments @("status", "--porcelain")
  $statusTarget = Invoke-Git -GitExe $gitExe -Arguments @("status", "--porcelain", "--", $relativeAlbumPath)

  $targetPaths = $statusTarget.Output | Where-Object { $_.Trim() } | ForEach-Object { Get-StatusPath -StatusLine $_ }
  if (-not $targetPaths) {
    Show-Info @"
No new photos were found to push.

This is good news:
- Saturday pictures will not duplicate.
- Sunday pictures will only be added when new files are there.
"@
    return
  }

  $outsideChanges = $statusAll.Output |
    Where-Object { $_.Trim() } |
    ForEach-Object { Get-StatusPath -StatusLine $_ } |
    Where-Object { $_ -and (-not $_.StartsWith($relativeAlbumPath, [System.StringComparison]::OrdinalIgnoreCase)) }

  if ($outsideChanges) {
    throw "There are other changes in the website folder right now. Please stop here and ask for help before auto-publishing."
  }

  Invoke-Git -GitExe $gitExe -Arguments @("add", "--", $relativeAlbumPath) | Out-Null
  Invoke-Git -GitExe $gitExe -Arguments @("commit", "-m", "Update photo album: $albumFolderName") | Out-Null
  Invoke-Git -GitExe $gitExe -Arguments @("push", "origin", "main") | Out-Null

  $summaryLines = @(
    "Finished photo export.",
    "Processed:",
    "Resized:",
    "Copied:",
    "Input size:",
    "Output size:",
    "Saved:"
  )

  $summary = $output |
    ForEach-Object { $_.ToString() } |
    Where-Object {
      $line = $_
      $summaryLines | Where-Object { $line.StartsWith($_, [System.StringComparison]::OrdinalIgnoreCase) }
    }

  $message = @"
Weekend album published.

Album:
$cleanAlbumTitle

Website folder:
$albumFolderName

What happens next time:
- Put new Sunday photos in the same Sunday folder.
- Run this same publisher again.
- Old Saturday files stay skipped.
- Only the new files get pushed.
"@

  if ($summary) {
    $message += "`r`nSummary:`r`n" + ($summary -join "`r`n")
  }

  Show-Info $message
} catch {
  Show-ErrorBox $_.Exception.Message
  exit 1
}
