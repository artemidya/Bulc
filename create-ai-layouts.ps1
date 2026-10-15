[CmdletBinding()]
param(
    [string]$WorkbookPath = "",
    [string]$SourceFolder = "",
    [string]$OutputFolder = "",
    [double]$MaxRowWidthMm = 1000,
    [double]$GapMm = 20,
    [double]$SafeMaxArtboardWidthMm = 205,
    [double]$SafeMaxArtboardHeightMm = 115,
    [switch]$PrepareOnly,
    [switch]$RepeatByItog
)

$ErrorActionPreference = "Stop"

$Root = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath = Join-Path $Root "List.xlsx" }
if ([string]::IsNullOrWhiteSpace($SourceFolder)) { $SourceFolder = Join-Path $Root "M" }
if ([string]::IsNullOrWhiteSpace($OutputFolder)) { $OutputFolder = Join-Path $Root "OUT" }

$BuildFolder = Join-Path $Root "build"
$AliasPath = Join-Path $Root "group-file-aliases.tsv"
$DataPath = Join-Path $BuildFolder "layout-data.json"
$PreflightPath = Join-Path $BuildFolder "layout-preflight.txt"
$ReportPath = Join-Path $BuildFolder "layout-report.txt"
$SnapshotPath = Join-Path $BuildFolder "List.snapshot.xlsx"
$JsxPath = Join-Path $Root "scripts\compose-ai-layouts.jsx"

New-Item -ItemType Directory -Force -Path $BuildFolder | Out-Null

function ConvertTo-ColumnNumber {
    param([string]$Letters)

    $n = 0
    foreach ($ch in $Letters.ToUpperInvariant().ToCharArray()) {
        $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $n
}

function Get-CellValue {
    param(
        [xml]$SheetXml,
        [object]$Cell,
        [string[]]$SharedStrings
    )

    $type = $Cell.GetAttribute("t")
    $valueNode = $Cell.SelectSingleNode("*[local-name()='v']")
    $value = ""
    if ($valueNode -ne $null) {
        $value = [string]$valueNode.InnerText
    }

    if ($type -eq "s") {
        if ($value -ne "") {
            return $SharedStrings[[int]$value]
        }
        return ""
    }

    if ($type -eq "inlineStr") {
        $parts = @()
        foreach ($textNode in $Cell.SelectNodes(".//*[local-name()='t']")) {
            $parts += [string]$textNode.InnerText
        }
        return ($parts -join "")
    }

    return $value
}

function Read-ZipEntryText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$Name
    )

    $entry = $Zip.GetEntry($Name)
    if ($entry -eq $null) {
        return $null
    }

    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    try {
        return $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }
}

function Join-XlsxPath {
    param(
        [string]$BasePath,
        [string]$Target
    )

    $targetPath = $Target -replace "\\", "/"
    if ($targetPath.StartsWith("/")) {
        return $targetPath.TrimStart("/")
    }

    $baseDir = Split-Path -Parent $BasePath
    $combined = Join-Path $baseDir $targetPath
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($part in ($combined -replace "\\", "/" -split "/")) {
        if ($part -eq "" -or $part -eq ".") { continue }
        if ($part -eq "..") {
            if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) }
            continue
        }
        $parts.Add($part)
    }
    return ($parts -join "/")
}

function Read-XlsxRows {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $sharedStrings = @()
        $sharedText = Read-ZipEntryText -Zip $zip -Name "xl/sharedStrings.xml"
        if ($sharedText) {
            [xml]$sharedXml = $sharedText
            foreach ($item in $sharedXml.SelectNodes("//*[local-name()='si']")) {
                $parts = @()
                foreach ($textNode in $item.SelectNodes(".//*[local-name()='t']")) {
                    $parts += [string]$textNode.InnerText
                }
                $sharedStrings += ($parts -join "")
            }
        }

        [xml]$workbookXml = Read-ZipEntryText -Zip $zip -Name "xl/workbook.xml"
        [xml]$relsXml = Read-ZipEntryText -Zip $zip -Name "xl/_rels/workbook.xml.rels"
        $firstSheet = $workbookXml.SelectSingleNode("//*[local-name()='sheet']")
        if ($firstSheet -eq $null) {
            throw "Workbook has no sheets."
        }

        $relId = $firstSheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        if ([string]::IsNullOrWhiteSpace($relId)) {
            $relId = $firstSheet.GetAttribute("r:id")
        }

        $relationship = $relsXml.SelectSingleNode("//*[local-name()='Relationship' and @Id='$relId']")
        if ($relationship -eq $null) {
            throw "Cannot find worksheet relationship $relId."
        }

        $sheetPath = Join-XlsxPath -BasePath "xl/workbook.xml" -Target ([string]$relationship.Target)
        [xml]$sheetXml = Read-ZipEntryText -Zip $zip -Name $sheetPath

        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($rowNode in $sheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
            $cells = @{}
            foreach ($cell in $rowNode.SelectNodes("*[local-name()='c']")) {
                $ref = $cell.GetAttribute("r")
                $letters = ([regex]::Match($ref, "^[A-Z]+")).Value
                if ([string]::IsNullOrWhiteSpace($letters)) { continue }
                $column = ConvertTo-ColumnNumber $letters
                $cells[$column] = Get-CellValue -SheetXml $sheetXml -Cell $cell -SharedStrings $sharedStrings
            }

            $rowNumber = 0
            [void][int]::TryParse($rowNode.GetAttribute("r"), [ref]$rowNumber)
            $rows.Add([pscustomobject]@{
                RowNumber = $rowNumber
                Cells = $cells
            })
        }

        return $rows
    }
    finally {
        $zip.Dispose()
    }
}

function Get-TableValue {
    param(
        [object]$Row,
        [int]$Column
    )

    if ($Row.Cells.ContainsKey($Column)) {
        return ([string]$Row.Cells[$Column]).Trim()
    }
    return ""
}

function Test-IncludedTotal {
    param([string]$Value)

    $text = ([string]$Value).Trim()
    if ($text.Length -eq 0) { return $false }

    $numberText = $text.Replace(",", ".")
    $number = 0.0
    $ok = [double]::TryParse(
        $numberText,
        [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    )

    if ($ok) { return $number -ne 0 }
    return $true
}

function Get-RepeatCount {
    param([string]$Value)

    if (-not $RepeatByItog) { return 1 }

    $numberText = ([string]$Value).Trim().Replace(",", ".")
    $number = 0.0
    $ok = [double]::TryParse(
        $numberText,
        [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    )

    if (-not $ok) { return 1 }
    $count = [int][math]::Floor($number)
    if ($count -lt 1) { return 1 }
    return $count
}

function ConvertTo-SafeFileName {
    param([string]$Name)

    $safe = ([string]$Name).Trim()
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$ch, "_")
    }
    $safe = ($safe -replace "\s+", " ").Trim(" ", ".")
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = "output" }
    return $safe
}

function ConvertTo-NormalKey {
    param([string]$Text)

    return (([string]$Text).ToLowerInvariant() -replace "[\s\p{P}\p{S}_-]+", "")
}

function ConvertTo-WordKey {
    param([string]$Text)

    $words = ([string]$Text).ToLowerInvariant() -split "[\s\p{P}\p{S}_-]+" | Where-Object { $_ -ne "" }
    return (($words | Sort-Object) -join "|")
}

function Read-GroupAliases {
    param([string]$Path)

    $aliases = @{}
    if (-not (Test-Path -LiteralPath $Path)) {
        return $aliases
    }

    foreach ($line in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $clean = ([string]$line).Trim()
        $clean = $clean.TrimStart([char]0xFEFF)
        if ($clean.Length -eq 0 -or $clean.StartsWith("#")) { continue }

        $parts = $line -split "`t"
        if ($parts.Count -lt 2) { continue }

        $group = ([string]$parts[0]).Trim().TrimStart([char]0xFEFF)
        $file = ([string]$parts[1]).Trim()
        if ($group.Length -gt 0 -and $file.Length -gt 0) {
            $aliases[$group] = $file
        }
    }

    return $aliases
}

function Resolve-AiFile {
    param(
        [string]$Group,
        [hashtable]$Aliases,
        [hashtable]$ByName,
        [hashtable]$ByBase,
        [hashtable]$ByNormal,
        [hashtable]$ByWords,
        [string]$SourceFolderPath
    )

    $candidates = New-Object System.Collections.Generic.List[string]
    if ($Aliases.ContainsKey($Group)) {
        $candidates.Add([string]$Aliases[$Group])
    }

    $candidates.Add($Group)
    $candidates.Add("$Group.ai")

    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }

        if ([System.IO.Path]::IsPathRooted($candidate) -and (Test-Path -LiteralPath $candidate)) {
            return (Get-Item -LiteralPath $candidate)
        }

        $name = [System.IO.Path]::GetFileName($candidate)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($candidate)

        if ($ByName.ContainsKey($name)) { return $ByName[$name] }
        if ($ByBase.ContainsKey($base)) { return $ByBase[$base] }

        $combined = Join-Path $SourceFolderPath $name
        if (Test-Path -LiteralPath $combined) {
            return (Get-Item -LiteralPath $combined)
        }
    }

    $normalKey = ConvertTo-NormalKey $Group
    if ($ByNormal.ContainsKey($normalKey)) { return $ByNormal[$normalKey] }

    $wordKey = ConvertTo-WordKey $Group
    if ($ByWords.ContainsKey($wordKey)) { return $ByWords[$wordKey] }

    return $null
}

$WorkbookFullPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$SourceFolderFullPath = (Resolve-Path -LiteralPath $SourceFolder).Path
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$OutputFolderFullPath = (Resolve-Path -LiteralPath $OutputFolder).Path

Copy-Item -LiteralPath $WorkbookFullPath -Destination $SnapshotPath -Force

$rows = Read-XlsxRows -Path $SnapshotPath
if ($rows.Count -lt 2) {
    throw "Workbook does not contain data rows."
}

$headerRow = $rows | Where-Object { $_.Cells.Count -gt 0 } | Select-Object -First 1
if ($headerRow -eq $null) {
    throw "Workbook header row is empty."
}

$dataRows = $rows | Where-Object { $_.RowNumber -gt $headerRow.RowNumber }

$ColumnMap = [ordered]@{
    Address = 1
    Group = 2
    Product = 3
    Total = 4
}

$aiFiles = Get-ChildItem -LiteralPath $SourceFolderFullPath -Filter "*.ai" -File
$byName = @{}
$byBase = @{}
$byNormal = @{}
$byWords = @{}
foreach ($file in $aiFiles) {
    $byName[$file.Name] = $file
    $byBase[$file.BaseName] = $file
    $byNormal[(ConvertTo-NormalKey $file.BaseName)] = $file
    $wordKey = ConvertTo-WordKey $file.BaseName
    if (-not $byWords.ContainsKey($wordKey)) {
        $byWords[$wordKey] = $file
    }
}

$aliases = Read-GroupAliases -Path $AliasPath
$items = New-Object System.Collections.Generic.List[object]
$warnings = New-Object System.Collections.Generic.List[string]
$outputNameByAddress = @{}
$usedOutputNames = @{}

foreach ($row in $dataRows) {
    $address = Get-TableValue -Row $row -Column $ColumnMap.Address
    $group = Get-TableValue -Row $row -Column $ColumnMap.Group
    $product = Get-TableValue -Row $row -Column $ColumnMap.Product
    $total = Get-TableValue -Row $row -Column $ColumnMap.Total

    if ($address -eq "" -and $group -eq "" -and $product -eq "" -and $total -eq "") {
        continue
    }
    if (-not (Test-IncludedTotal $total)) {
        continue
    }
    if ($address -eq "" -or $group -eq "" -or $product -eq "") {
        $warnings.Add("Row $($row.RowNumber): skipped because address, group, or product is empty.")
        continue
    }

    if (-not $outputNameByAddress.ContainsKey($address)) {
        $baseName = ConvertTo-SafeFileName $address
        $fileName = "$baseName.ai"
        $suffix = 2
        while ($usedOutputNames.ContainsKey($fileName)) {
            $fileName = "$baseName`_$suffix.ai"
            $suffix++
        }
        $outputNameByAddress[$address] = $fileName
        $usedOutputNames[$fileName] = $address
    }

    $sourceFile = Resolve-AiFile `
        -Group $group `
        -Aliases $aliases `
        -ByName $byName `
        -ByBase $byBase `
        -ByNormal $byNormal `
        -ByWords $byWords `
        -SourceFolderPath $SourceFolderFullPath

    $repeatCount = Get-RepeatCount $total
    for ($copy = 1; $copy -le $repeatCount; $copy++) {
        $items.Add([pscustomobject]@{
            address = $address
            group = $group
            product = $product
            total = $total
            sourceFile = if ($sourceFile -ne $null) { $sourceFile.FullName } else { $null }
            sourceFileName = if ($sourceFile -ne $null) { $sourceFile.Name } else { $null }
            outputFileName = $outputNameByAddress[$address]
            rowNumber = $row.RowNumber
            copyIndex = $copy
        })
    }
}

$data = [ordered]@{
    generatedAt = (Get-Date).ToString("s")
    settings = [ordered]@{
        workbookPath = $WorkbookFullPath
        sourceFolder = $SourceFolderFullPath
        outputFolder = $OutputFolderFullPath
        maxRowWidthMm = $MaxRowWidthMm
        gapMm = $GapMm
        safeMaxArtboardWidthMm = $SafeMaxArtboardWidthMm
        safeMaxArtboardHeightMm = $SafeMaxArtboardHeightMm
        repeatByItog = [bool]$RepeatByItog
        reportPath = $ReportPath
    }
    items = $items
}

$json = $data | ConvertTo-Json -Depth 8
Set-Content -LiteralPath $DataPath -Value $json -Encoding UTF8

$preflight = New-Object System.Collections.Generic.List[string]
$preflight.Add("Generated: $((Get-Date).ToString('s'))")
$preflight.Add("Workbook: $WorkbookFullPath")
$preflight.Add("Source folder: $SourceFolderFullPath")
$preflight.Add("Output folder: $OutputFolderFullPath")
$preflight.Add("Data file: $DataPath")
$preflight.Add("Layout: max row width ${MaxRowWidthMm}mm, gap ${GapMm}mm, safe cell ${SafeMaxArtboardWidthMm}x${SafeMaxArtboardHeightMm}mm")
$preflight.Add("")
$preflight.Add("Column mapping: A=Address, B=Group, C=Product, D=Total")
$headerPreview = 1..4 | ForEach-Object { Get-TableValue -Row $headerRow -Column $_ }
$preflight.Add("Detected headers A-D: " + ($headerPreview -join " | "))
$preflight.Add("")
$preflight.Add("Items prepared: $($items.Count)")
$addressGroups = $items | Group-Object address | Sort-Object Name
$preflight.Add("Output files: $($addressGroups.Count)")
foreach ($addressGroup in $addressGroups) {
    $first = $addressGroup.Group | Select-Object -First 1
    $preflight.Add("  $($addressGroup.Name): $($addressGroup.Count) item(s) -> $($first.outputFileName)")
}

$preflight.Add("")
$preflight.Add("Groups:")
foreach ($groupInfo in ($items | Group-Object group | Sort-Object Name)) {
    $first = $groupInfo.Group | Select-Object -First 1
    $status = if ($first.sourceFileName) { $first.sourceFileName } else { "MISSING AI FILE" }
    $preflight.Add("  $($groupInfo.Name): $($groupInfo.Count) item(s) -> $status")
}

$missingSources = $items | Where-Object { $_.sourceFile -eq $null }
if ($missingSources.Count -gt 0) {
    $preflight.Add("")
    $preflight.Add("Missing source files:")
    foreach ($missing in $missingSources) {
        $preflight.Add("  Row $($missing.rowNumber): group '$($missing.group)', product '$($missing.product)'")
    }
}

if ($warnings.Count -gt 0) {
    $preflight.Add("")
    $preflight.Add("Warnings:")
    foreach ($warning in $warnings) {
        $preflight.Add("  $warning")
    }
}

Set-Content -LiteralPath $PreflightPath -Value $preflight -Encoding UTF8

Write-Host "Prepared $($items.Count) item(s) for $($addressGroups.Count) output file(s)."
Write-Host "Data: $DataPath"
Write-Host "Preflight: $PreflightPath"

if ($PrepareOnly) {
    Write-Host "PrepareOnly is set; Illustrator was not started."
    return
}

if (-not (Test-Path -LiteralPath $JsxPath)) {
    throw "Missing Illustrator script: $JsxPath"
}

try {
    $illustrator = New-Object -ComObject Illustrator.Application
    [void]$illustrator.DoJavaScriptFile((Resolve-Path -LiteralPath $JsxPath).Path)
    Write-Host "Illustrator script started. Report: $ReportPath"
}
catch {
    Write-Warning "Could not start Illustrator through COM automation."
    Write-Warning $_.Exception.Message
    Write-Host "Open Illustrator and run this script manually:"
    Write-Host $JsxPath
    throw
}
