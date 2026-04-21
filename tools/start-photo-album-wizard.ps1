[CmdletBinding()]
param(
  [string]$SourceFolder,
  [string]$AlbumName,
  [string]$DestinationRoot = (Join-Path $PSScriptRoot "..\media\photo-albums"),
  [int]$MaxLongEdge = 2200,
  [ValidateRange(1, 100)]
  [int]$JpegQuality = 82,
  [switch]$Force,
  [switch]$SkipOpenFolder,
  [switch]$Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$windowTitle = "Photo Album Wizard"
$exportScript = Join-Path $PSScriptRoot "export-photo-albums.ps1"
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
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

  $message = @"
Album ready.

Saved here:
$destinationFolder

Next steps:
1. Open GitHub Desktop.
2. Make sure this repo says mud-motion-site.
3. In Changes, leave the boxes checked.
4. Type a short message.
5. Click Commit to main.
6. Click Push origin.
"@

  if ($summary) {
    $message += "`r`nSummary:`r`n" + ($summary -join "`r`n")
  }

  Show-Info $message
} catch {
  Show-ErrorBox $_.Exception.Message
  exit 1
}
