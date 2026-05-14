[CmdletBinding()]
param(
  [string]$SourceFolder,
  [string]$AlbumName,
  [string]$DestinationRoot = "",
  [int]$MaxLongEdge = 2200,
  [ValidateRange(1, 100)]
  [int]$JpegQuality = 82,
  [switch]$Force,
  [switch]$SkipGit,
  [switch]$SkipOpenFolder,
  [switch]$Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $PSCommandPath }
if (-not $DestinationRoot) {
  $DestinationRoot = Join-Path $scriptRoot "..\media\photo-albums"
}

$windowTitle = "Photo Album Wizard"
$exportScript = Join-Path $scriptRoot "export-photo-albums.ps1"
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptRoot ".."))
$destinationRootFull = [System.IO.Path]::GetFullPath($DestinationRoot)

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

function Ask-YesNo {
  param([string]$Message)

  if ($Quiet) {
    throw "Quiet mode cannot ask yes or no questions. Use -Force if you want to replace files in an existing album."
  }

  $result = [System.Windows.Forms.MessageBox]::Show(
    $Message,
    $windowTitle,
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
  )

  return $result -eq [System.Windows.Forms.DialogResult]::Yes
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

  $cleanName = $builder.ToString() -replace "\s+", " "
  $cleanName = $cleanName.Trim().TrimEnd(".")
  return $cleanName
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

  throw "Git was not found. Please install GitHub Desktop or Git on this computer first."
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

try {
  if (-not (Test-Path -LiteralPath $exportScript)) {
    throw "Could not find the export tool at $exportScript"
  }

  if (-not (Test-Path -LiteralPath $destinationRootFull)) {
    New-Item -ItemType Directory -Path $destinationRootFull -Force | Out-Null
  }

  if (-not $SourceFolder) {
    $SourceFolder = Select-Folder -Description "Pick the folder that has the full-size photos for one album." -SelectedPath $env:USERPROFILE
    if (-not $SourceFolder) {
      return
    }
  }

  $sourceFolderFull = [System.IO.Path]::GetFullPath($SourceFolder)
  if (-not (Test-Path -LiteralPath $sourceFolderFull)) {
    throw "Source folder not found: $sourceFolderFull"
  }

  if ($sourceFolderFull.StartsWith($repoRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Please pick the original photos folder outside the website folder. The wizard makes a web-sized copy inside the website."
  }

  $defaultAlbumName = Split-Path -Leaf $sourceFolderFull
  if (-not $AlbumName) {
    $AlbumName = [Microsoft.VisualBasic.Interaction]::InputBox(
      "Type the album name exactly how it should show on the website.",
      $windowTitle,
      $defaultAlbumName
    )

    if (-not $AlbumName) {
      return
    }
  }

  $cleanAlbumName = Normalize-AlbumName -Name $AlbumName
  if (-not $cleanAlbumName) {
    throw "Album name is blank after cleaning it up. Please try again."
  }

  if ($cleanAlbumName -ne $AlbumName) {
    Show-Info "The album folder name was cleaned up a little to remove characters Windows does not allow.`r`n`r`nNew folder name:`r`n$cleanAlbumName"
  }

  $destinationFolder = Join-Path $destinationRootFull $cleanAlbumName
  $relativeAlbumPath = Get-RelativeRepoPath -Path $destinationFolder
  $destinationExists = Test-Path -LiteralPath $destinationFolder
  if ($destinationExists -and (-not $Force)) {
    $Force = Ask-YesNo "An album folder with this name already exists.`r`n`r`nClick Yes to replace same-name files in that album.`r`nClick No to cancel."
    if (-not $Force) {
      return
    }
  }

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

  if (-not $SkipOpenFolder) {
    Start-Process explorer.exe $destinationFolder
  }

  if ($SkipGit) {
    $message = @"
Album ready.

Saved here:
$destinationFolder

Git push was skipped because SkipGit was turned on.
"@

    if ($summary) {
      $message += "`r`nSummary:`r`n" + ($summary -join "`r`n")
    }

    Show-Info $message
    return
  }

  $gitExe = Get-GitCommand
  $statusAll = Invoke-Git -GitExe $gitExe -Arguments @("status", "--porcelain")
  $statusTarget = Invoke-Git -GitExe $gitExe -Arguments @("status", "--porcelain", "--", $relativeAlbumPath)

  $targetPaths = $statusTarget.Output | Where-Object { $_.Trim() } | ForEach-Object { Get-StatusPath -StatusLine $_ }
  if (-not $targetPaths) {
    Show-Info @"
No new photos were found to push.

This usually means the album was already uploaded.
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
  Invoke-Git -GitExe $gitExe -Arguments @("commit", "-m", "Add photo album: $cleanAlbumName") | Out-Null
  Invoke-Git -GitExe $gitExe -Arguments @("push", "origin", "main") | Out-Null

  $message = @"
Album ready and pushed.

Saved here:
$destinationFolder

GitHub and Netlify should update the website shortly.
"@

  if ($summary) {
    $message += "`r`nSummary:`r`n" + ($summary -join "`r`n")
  }

  Show-Info $message
} catch {
  Show-ErrorBox $_.Exception.Message
  exit 1
}
