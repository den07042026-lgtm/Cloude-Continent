param(
    [string]$RepositoryRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$CardArchive = 'C:\Users\Admin\Desktop\Итого\img',
    [string]$SourceArchive = 'C:\Users\Admin\Desktop\На сортировку 08.05\images'
)

$ErrorActionPreference = 'Stop'
$archiveRoot = Join-Path $RepositoryRoot 'product-assets\archive-2026-08'
$cardRoot = Join-Path $archiveRoot 'cards'
$sourceRoot = Join-Path $archiveRoot 'sources'
New-Item -ItemType Directory -Force -Path $cardRoot, $sourceRoot | Out-Null

function Normalize-Key([string]$value) {
    return $value.Trim().ToLowerInvariant()
}

function Add-HardLink([string]$source, [string]$target) {
    $parent = Split-Path -Parent $target
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }
    if (Test-Path -LiteralPath $target) {
        if ((Get-Item -LiteralPath $target).Length -ne (Get-Item -LiteralPath $source).Length) {
            throw "Conflicting target: $target"
        }
        return
    }
    New-Item -ItemType HardLink -Path $target -Target $source | Out-Null
}

$skuCategories = @{}
$cardCount = 0
foreach ($file in Get-ChildItem -LiteralPath $CardArchive -File -Recurse) {
    if ($file.BaseName -notmatch '^(?<stem>.+)_(?<role>compat|info)$') {
        throw "Unexpected card name: $($file.FullName)"
    }
    $role = $Matches.role
    $stem = $Matches.stem
    if ($stem -notmatch '^(?<sku>.+?)(?:_rsk)?_(?<sequence>[0-9]+)$') {
        throw "Unexpected card sequence: $($file.FullName)"
    }

    $sku = Normalize-Key $Matches.sku
    $sequence = [int]$Matches.sequence
    $category = Normalize-Key (Split-Path $file.DirectoryName -Leaf)
    if (-not $skuCategories.ContainsKey($sku)) { $skuCategories[$sku] = $category }
    $target = Join-Path $cardRoot "$category\$sku\$('{0:D2}' -f $sequence)-$role$($file.Extension.ToLowerInvariant())"
    Add-HardLink $file.FullName $target
    $cardCount++
}

$sourceCount = 0
foreach ($file in Get-ChildItem -LiteralPath $SourceArchive -File -Recurse) {
    if ($file.BaseName -notmatch '^(?<sku>.+?)(?:_rsk)?_(?<sequence>[0-9]+)$') {
        throw "Unexpected source name: $($file.FullName)"
    }
    $sku = Normalize-Key $Matches.sku
    $sequence = [int]$Matches.sequence
    $target = Join-Path $sourceRoot "$sku\$('{0:D2}' -f $sequence)$($file.Extension.ToLowerInvariant())"
    Add-HardLink $file.FullName $target
    $sourceCount++
}

$manifest = foreach ($file in Get-ChildItem -LiteralPath $archiveRoot -File -Recurse | Sort-Object FullName) {
    $relative = [IO.Path]::GetRelativePath($RepositoryRoot, $file.FullName).Replace('\', '/')
    $parts = $relative.Split('/')
    if ($parts[2] -eq 'cards') {
        $category = $parts[3]
        $sku = $parts[4]
        $role = if ($file.BaseName -match '-(?<role>compat|info)$') { $Matches.role } else { 'card' }
    }
    else {
        $category = if ($skuCategories.ContainsKey($parts[3])) { $skuCategories[$parts[3]] } else { '' }
        $sku = $parts[3]
        $role = 'source'
    }
    [PSCustomObject]@{
        sku = $sku
        category = $category
        role = $role
        path = $relative
        bytes = $file.Length
    }
}
$manifest | Export-Csv -LiteralPath (Join-Path $archiveRoot 'manifest.csv') -NoTypeInformation -Encoding utf8
Write-Host "Imported $cardCount cards and $sourceCount source photos; indexed $($manifest.Count) files."
