param(
    [string]$RepositoryRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$SourcePhotoRoot = 'C:\Users\Admin\Documents\Codex\2026-07-20\vyt\work\mikado_growth_full\images'
)

$ErrorActionPreference = 'Stop'

$assetRoot = Join-Path $RepositoryRoot 'product-assets'
$cardsRoot = Join-Path $assetRoot 'cards'
$sourcesRoot = Join-Path $assetRoot 'sources'
$sharedRoot = Join-Path $assetRoot 'shared'
$legacyRoots = @(
    (Join-Path $RepositoryRoot 'ozon-images\catalog'),
    (Join-Path $RepositoryRoot 'ozon-images\next225-2026-08-07'),
    (Join-Path $RepositoryRoot 'ozon-images\test-batch')
)

New-Item -ItemType Directory -Force -Path $cardsRoot, $sourcesRoot, $sharedRoot | Out-Null

$seen = @{}
foreach ($legacyRoot in $legacyRoots) {
    if (-not (Test-Path -LiteralPath $legacyRoot)) { continue }

    foreach ($file in Get-ChildItem -LiteralPath $legacyRoot -File) {
        $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash

        if ($file.Name -match '^(?<sku>.+)_(?<role>compat|info)\.jpg$') {
            $sku = $Matches.sku.ToLowerInvariant()
            $role = if ($Matches.role -eq 'compat') { 'compatibility' } else { 'info' }
            $targetDir = Join-Path $cardsRoot $sku
            $target = Join-Path $targetDir "$role.jpg"
        }
        else {
            $targetDir = $sharedRoot
            $target = Join-Path $targetDir $file.Name.ToLowerInvariant()
        }

        New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
        if ($seen.ContainsKey($target)) {
            if ($seen[$target] -ne $hash) {
                throw "Conflicting files target the same path: $target"
            }
            Remove-Item -LiteralPath $file.FullName
            continue
        }

        Move-Item -LiteralPath $file.FullName -Destination $target
        $seen[$target] = $hash
    }
}

if (Test-Path -LiteralPath $SourcePhotoRoot) {
    foreach ($file in Get-ChildItem -LiteralPath $SourcePhotoRoot -File) {
        if ($file.BaseName -notmatch '^(?<sku>.+)_(?<sequence>[0-9]+)$') {
            throw "Unexpected source-photo name: $($file.Name)"
        }

        $sku = $Matches.sku.ToLowerInvariant()
        $sequence = [int]$Matches.sequence
        $targetDir = Join-Path $sourcesRoot $sku
        $target = Join-Path $targetDir ("{0:D2}{1}" -f $sequence, $file.Extension.ToLowerInvariant())
        New-Item -ItemType Directory -Force -Path $targetDir | Out-Null

        if (Test-Path -LiteralPath $target) {
            $sourceHash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
            $targetHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
            if ($sourceHash -ne $targetHash) { throw "Conflicting source photo: $target" }
        }
        else {
            Copy-Item -LiteralPath $file.FullName -Destination $target
        }
    }
}

foreach ($legacyRoot in $legacyRoots) {
    if ((Test-Path -LiteralPath $legacyRoot) -and -not (Get-ChildItem -LiteralPath $legacyRoot -Force)) {
        Remove-Item -LiteralPath $legacyRoot
    }
}
$oldRoot = Join-Path $RepositoryRoot 'ozon-images'
if ((Test-Path -LiteralPath $oldRoot) -and -not (Get-ChildItem -LiteralPath $oldRoot -Force)) {
    Remove-Item -LiteralPath $oldRoot
}

Add-Type -AssemblyName System.Drawing
$imageExtensions = @('.jpg', '.jpeg', '.png', '.webp', '.gif', '.avif', '.tif', '.tiff', '.heic')
$rows = foreach ($file in Get-ChildItem -LiteralPath $assetRoot -File -Recurse |
    Where-Object { $imageExtensions -contains $_.Extension.ToLowerInvariant() } |
    Sort-Object FullName) {
    $relative = [IO.Path]::GetRelativePath($RepositoryRoot, $file.FullName).Replace('\', '/')
    $parts = $relative.Split('/')
    $category = $parts[1]
    $sku = if ($category -eq 'shared') { '' } else { $parts[2] }
    $role = if ($category -eq 'cards') { $file.BaseName } elseif ($category -eq 'sources') { 'source' } else { 'shared' }

    $width = $null
    $height = $null
    try {
        $image = [System.Drawing.Image]::FromFile($file.FullName)
        $width = $image.Width
        $height = $image.Height
        $image.Dispose()
    }
    catch { }

    [PSCustomObject]@{
        sku = $sku
        category = $category
        role = $role
        path = $relative
        width = $width
        height = $height
        bytes = $file.Length
        sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    }
}

$rows | Export-Csv -LiteralPath (Join-Path $assetRoot 'manifest.csv') -NoTypeInformation -Encoding utf8
Write-Host "Indexed $($rows.Count) product assets."
