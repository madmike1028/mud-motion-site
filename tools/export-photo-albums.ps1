[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [string]$SourceRoot = "",
  [string]$DestinationRoot = "",
  [int]$MaxLongEdge = 2400,
  [ValidateRange(1, 100)]
  [int]$JpegQuality = 82,
  [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $PSCommandPath }
if (-not $SourceRoot) {
  $SourceRoot = Join-Path $scriptRoot "..\incoming-photo-albums"
}
if (-not $DestinationRoot) {
  $DestinationRoot = Join-Path $scriptRoot "..\media\photo-albums"
}

$supportedExtensions = @(".jpg", ".jpeg", ".png")
$sourceRootFull = [System.IO.Path]::GetFullPath($SourceRoot)
$destinationRootFull = [System.IO.Path]::GetFullPath($DestinationRoot)

if (-not (Test-Path -LiteralPath $sourceRootFull)) {
  throw "Source folder not found: $sourceRootFull"
}

if ($sourceRootFull -eq $destinationRootFull) {
  throw "Source and destination must be different folders. Export into media\photo-albums from a separate originals folder."
}

function Get-JpegEncoder {
  return [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
    Where-Object { $_.MimeType -eq "image/jpeg" } |
    Select-Object -First 1
}

function Get-ResizeDimensions {
  param(
    [int]$Width,
    [int]$Height,
    [int]$MaxEdge
  )

  $longEdge = [Math]::Max($Width, $Height)
  if ($longEdge -le $MaxEdge) {
    return [pscustomobject]@{
      Width = $Width
      Height = $Height
      Resized = $false
    }
  }

  $scale = $MaxEdge / [double]$longEdge
  return [pscustomobject]@{
    Width = [Math]::Max(1, [int][Math]::Round($Width * $scale))
    Height = [Math]::Max(1, [int][Math]::Round($Height * $scale))
    Resized = $true
  }
}

function Get-RelativePathText {
  param(
    [string]$BasePath,
    [string]$TargetPath
  )

  $baseFullPath = [System.IO.Path]::GetFullPath($BasePath).TrimEnd("\")
  $targetFullPath = [System.IO.Path]::GetFullPath($TargetPath)
  $baseUri = New-Object System.Uri($baseFullPath + "\")
  $targetUri = New-Object System.Uri($targetFullPath)
  $relativeUri = $baseUri.MakeRelativeUri($targetUri)
  return [System.Uri]::UnescapeDataString($relativeUri.ToString()).Replace("/", "\")
}

function Set-ImageOrientation {
  param(
    [System.Drawing.Image]$Image
  )

  $orientationId = 0x0112
  if (-not ($Image.PropertyIdList -contains $orientationId)) {
    return
  }

  $orientation = [BitConverter]::ToUInt16($Image.GetPropertyItem($orientationId).Value, 0)
  switch ($orientation) {
    2 { $Image.RotateFlip([System.Drawing.RotateFlipType]::RotateNoneFlipX) }
    3 { $Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate180FlipNone) }
    4 { $Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate180FlipX) }
    5 { $Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipX) }
    6 { $Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipNone) }
    7 { $Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipX) }
    8 { $Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipNone) }
    default { }
  }

  try {
    $Image.RemovePropertyItem($orientationId)
  } catch {
  }
}

function Save-ImageFile {
  param(
    [System.Drawing.Image]$Image,
    [string]$DestinationPath,
    [int]$Quality
  )

  $extension = [System.IO.Path]::GetExtension($DestinationPath).ToLowerInvariant()
  if ($extension -in @(".jpg", ".jpeg")) {
    $encoder = Get-JpegEncoder
    $encoderParameters = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encoderParameters.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter(
      [System.Drawing.Imaging.Encoder]::Quality,
      [long]$Quality
    )
    try {
      $Image.Save($DestinationPath, $encoder, $encoderParameters)
    } finally {
      $encoderParameters.Dispose()
    }
    return
  }

  if ($extension -eq ".png") {
    $Image.Save($DestinationPath, [System.Drawing.Imaging.ImageFormat]::Png)
    return
  }

  throw "Unsupported output format: $DestinationPath"
}

if (-not (Test-Path -LiteralPath $destinationRootFull)) {
  New-Item -ItemType Directory -Path $destinationRootFull | Out-Null
}

$files = Get-ChildItem -LiteralPath $sourceRootFull -Recurse -File |
  Where-Object {
    $extension = $_.Extension.ToLowerInvariant()
    $fullPath = [System.IO.Path]::GetFullPath($_.FullName)
    ($supportedExtensions -contains $extension) -and
    (-not $fullPath.StartsWith($destinationRootFull, [System.StringComparison]::OrdinalIgnoreCase))
  }

if (-not $files) {
  Write-Output "No supported image files found under $sourceRootFull"
  return
}

$processedCount = 0
$copiedCount = 0
$resizedCount = 0
$originalBytes = [int64]0
$outputBytes = [int64]0

foreach ($file in $files) {
  $relativePath = Get-RelativePathText -BasePath $sourceRootFull -TargetPath $file.FullName
  $destinationPath = Join-Path $destinationRootFull $relativePath
  $destinationDir = Split-Path -Parent $destinationPath

  if (-not (Test-Path -LiteralPath $destinationDir)) {
    New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
  }

  $processedCount++
  $originalBytes += $file.Length

  if ((Test-Path -LiteralPath $destinationPath) -and (-not $Force)) {
    Write-Output "Skipping existing file: $relativePath"
    $outputBytes += (Get-Item -LiteralPath $destinationPath).Length
    continue
  }

  if (-not $PSCmdlet.ShouldProcess($destinationPath, "Export optimized image")) {
    continue
  }

  $sourceImage = [System.Drawing.Image]::FromFile($file.FullName)
  try {
    Set-ImageOrientation -Image $sourceImage
    $size = Get-ResizeDimensions -Width $sourceImage.Width -Height $sourceImage.Height -MaxEdge $MaxLongEdge

    if (-not $size.Resized) {
      Copy-Item -LiteralPath $file.FullName -Destination $destinationPath -Force
      $copiedCount++
      $outputBytes += (Get-Item -LiteralPath $destinationPath).Length
      Write-Output "Copied original: $relativePath"
      continue
    }

    $bitmap = New-Object System.Drawing.Bitmap($size.Width, $size.Height)
    try {
      $horizontalResolution = if ($sourceImage.HorizontalResolution -gt 0) { $sourceImage.HorizontalResolution } else { 72 }
      $verticalResolution = if ($sourceImage.VerticalResolution -gt 0) { $sourceImage.VerticalResolution } else { 72 }
      $bitmap.SetResolution($horizontalResolution, $verticalResolution)

      $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
      try {
        $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $graphics.DrawImage($sourceImage, 0, 0, $size.Width, $size.Height)
      } finally {
        $graphics.Dispose()
      }

      Save-ImageFile -Image $bitmap -DestinationPath $destinationPath -Quality $JpegQuality
      $resizedCount++
      $outputBytes += (Get-Item -LiteralPath $destinationPath).Length
      Write-Output ("Resized: {0} -> {1}x{2}" -f $relativePath, $size.Width, $size.Height)
    } finally {
      $bitmap.Dispose()
    }
  } finally {
    $sourceImage.Dispose()
  }
}

$savedBytes = $originalBytes - $outputBytes
if ($savedBytes -lt 0) {
  $savedBytes = [int64]0
}
$savedMegabytes = [Math]::Round($savedBytes / 1MB, 2)
$originalMegabytes = [Math]::Round($originalBytes / 1MB, 2)
$outputMegabytes = [Math]::Round($outputBytes / 1MB, 2)

Write-Output ""
Write-Output "Finished photo export."
Write-Output "Processed: $processedCount"
Write-Output "Resized: $resizedCount"
Write-Output "Copied: $copiedCount"
Write-Output "Input size: $originalMegabytes MB"
Write-Output "Output size: $outputMegabytes MB"
Write-Output "Saved: $savedMegabytes MB"
