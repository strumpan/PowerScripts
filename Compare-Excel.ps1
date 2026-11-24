<#
.SYNOPSIS
    Compares two Excel files.

.DESCRIPTION
    Compares two Excel files using a unique key column.
    Column order does not affect the comparision.
    Results are exported to a CSV file.

.PARAMETER File1
    Original Excel file

.PARAMETER File2
    New Excel file

.PARAMETER KeyColumn  
    Column containing the unique record identifier

.PARAMETER Sheet
    Name of the Excel woriksheet to compare.

.PARAMETER Output
    Path to the output CSV file.

.PARAMETER Help
    Displays usage information.

.EXAMPLE
    .\Compare-Excel.ps1 -File1 Old.xlsx -File2 New.xlsx -KeyColumn Id

.EXAMPLE
    .\Compare-Excel.ps1 -Help

.NOTES
    Compatible with PowerShell 5.1.
    Designed for ConstrainedLanguage mode.
    Microsoft Excel is not required. 
#>

param(
    [string]$File1, 
    [string]$File2,
    [string]$KeyColumn,
    [string]$Sheet = "",
    [string]$Output = "$env:TEMP\ExcelDifferences.csv",
    [Alias('h','-help')]
    [switch]$Help
)

function Show-Help {
    Get-Help $PSCommandPath -Detailed
}


if ($Help -or $PSBoundParameters.Count -eq 0) {
    Show-Help
    return
}

if (-not $File1 -or
    -not $File2 -or
    -not $KeyColumn) {
    
    Write-Host "Missing required parameters."
    Show-Help
    return
}

$ErrorActionPreference = 'Stop'

Function Find-XmlMatches {
    param(
        [string]$Text, 
        [string]$Pattern
    )

    if ([string]::IsNullOrEmpty($Pattern)) {
        Get-PSCallStack | Select-Object Command, Location, ScriptName, ScriptLineNumber
        throw "Find-XmlMatches: Empty pattern"
    }

    if ([string]::IsNullOrEmpty($Text)) {
        return
    }

    $result = Select-String `
        -InputObject $Text `
        -Pattern $Pattern `
        -AllMatches

    if ($null -ne $result) {
        foreach ($matchItem in $result.Matches) {
                $matchItem
        }
    }
}

function Decode-XmlText {
    param([string]$Value)

    $Value = $Value.Replace('&lt;', '<').Replace('&gt;', '>')
    $Value = $Value.Replace('&quot;', '"').Replace('&apos;', "'")

    return $Value.Replace('&amp;', '&')
}

function Get-Attribute {
    param(
        [string]$Tag, 
        [string]$Name
    )

    $pattern = '(?:^|\s)' + $Name + '="([^"]*)"'

    if ($Tag -match $pattern) {
        return $Matches[1]
    }

    return ''
}

function Get-ColumnIndex {
    param([string]$Reference)

    if ($Reference -notmatch '^([A-Za-z]+)') {
        return 0
    }

    $letters = $Matches[1].ToUpperInvariant()
    $number = 0

    foreach ($letter in ($letters -split '')) {
        if ($letter -eq '') { continue }
        $alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        $position = $alphabet.IndexOf($letter) + 1
        $number = ($number * 26) + $position
    }

    return $number
}

function Format-ExcelDate {
    param(
        [string]$Value,
        [int]$FormatId,
        [bool]$Date1904

    )

    if ($Value -notmatch '^\d+([.,]\d+)?$') {
        return $Value
    } 

    $number = [double]$Value

    try {
        if ($FormatId -in @(18,19,20,21,45,47)) {
            $date = [datetime]::FromOADate($number)
            return $date.ToString('HH:mm:ss')
        }

        if ($Date1904) {
            $number += 1462
        }
        elseif ($number -lt 60) {
            $number += 1
        }
        elseif ($number -lt 61) {
            return $Value
        }

        $date = [datetime]::FromOADate($number)

        if ($FormatId -eq 22 -or
            $number %1 -ne 0) {
            return $date.ToString('yyyy-MM-dd HH:mm:ss')    
        }

        return $date.ToString('yyyy-MM-dd')
    }
    catch {
        return $Value
    }
}

function Get-ZipText {
    param(
        [string]$Path,
        [string]$Entry
    )

    $result = tar.exe -xOf $Path $Entry 2>$null

    if ($LASTEXITCODE -ne 0) {
        return ''
    }

    return ($result -join "`n")
}


function Read-ExcelSheet{
	param([string]$Path, [string]$SheetName, [string]$KeyName)

	$stylesXml = Get-ZipText $Path 'xl/styles.xml'

    $styleFormats = @{}

    if ($stylesXml -match '(?s)<(?:\w+:)?cellXfs\b[^>]*>(.*?)</(?:\w+:)?cellXfs>') {
        $xfXml = $Matches[1]
        $i = 0
        foreach ($xf in (Find-XmlMatches $xfXml '<(?:\w+:)?xf\b[^>]*>')) {
            $numFmtId = Get-Attribute $xf.Value 'numFmtId'
            $styleFormats[$i] = $numFmtId

            $i++
        }
    }

    foreach ($fmt in (Find-XmlMatches $stylesXml 'numFmt\b[^>]*/>')) {
        $id = Get-Attribute $fmt.Value 'numFmtId'
        $code = Get-Attribute $fmt.Value 'formatCode'
    }
	
    $workbook = Get-ZipText $Path 'xl/workbook.xml'
    $date1904 = $workbook -match 'date1904="(1|true)"'
    $relationships = Get-ZipText $Path 'xl/_rels/workbook.xml.rels'
    $sheetTags =( Find-XmlMatches $workbook '<(?:\w+:)?sheet\b[^>]*>')
    if ($sheetTags.Count -eq 0) { throw "Inga kalkylblad hittades i '$Path'." }

    $selectec = $null
    foreach ($s in $sheetTags) {
        $name = Decode-XmlText (Get-Attribute $s.Value 'name')
        if (($SheetName -eq '' -and $null -eq $selected) -or ($SheetName -ne '' -and $name -eq $SheetName)) {
            $selected = $s.Value
            $actualSheet = $name
            break
        }
    }
    if ($null -eq $selected) { throw "Bladet '$SheetName' saknas i '$Path'." }

    $rid = Get-Attribute $selected 'r:id'
    $target = ''
    foreach ($r in (Find-XmlMatches $relationships '<(?:\w+:)?Relationship\b[^>]*>')) {
        if ((Get-Attribute $r.Value 'Id') -eq $rid) {
            $target = Get-Attribute $r.Value 'Target'
            break
        }
    }
    if ($target -eq '') { throw "Kan inte hitta XML-filen för bladet '$actualSheet'." }
    $target = $target.Replace('\', '/')
    if ($target.StartsWith('/')) { $entry = $target.TrimStart('/') }
    elseif ($target.StartsWith('xl/')) { $entry = $target }
    else { $entry = 'xl/' + $target }
    $entry = $entry.Replace('xl/../', '')

    $sharedXml = Get-ZipText $Path 'xl/sharedStrings.xml' $true
    $shared = @()
    foreach ($si in (Find-XmlMatches $sharedXml '(?s)<(?:\w+:)?si(?:\s[^>]*)?>(.*?)</(?:\w+:)?si>')) {
        $value = ''
        foreach ($t in (Find-XmlMatches $si.Groups[1].Value '(?s)<(?:\w+:)?t(?:\s[^>]*)?>(.*?)</(?:\w+:)?t>')) {
            $value += Decode-XmlText $t.Groups[1].Value
        }
        $shared += $value
    }

    $xml = Get-ZipText $Path $entry
    $rows = @()
    foreach ($rowMatch in (Find-XmlMatches $xml '(?s)<(?:\w+:)?row\b[^>]*(?:/>|>.*?</(?:\w+:)?row>)')) {
        $rowXml = $rowMatch.Value
        $cells = @{}
        foreach ($cell in (Find-XmlMatches $rowXml '(?s)<(?:\w+:)?c\b[^>]*(?:/>|>.*?</(?:\w+:)?c>)')) {
            $cellXml = $cell.Value
            $openTag = (Find-XmlMatches $cellXml '^<[^>]+>').Value
            $ref = Get-Attribute $opentag 'r'
            $index = Get-ColumnIndex $ref
            $styleIndex = Get-Attribute $openTag 's'

            if ($index -eq 0) { continue }

            $type = Get-Attribute $openTag 't'
            $value = ''
            if ($type -eq 'inlineStr') {
                foreach ($t in (Find-XmlMatches $cellXml '(?s)<(?:\w+:)?t(?:\s[^>]*)?>(.*?)</(?:\w+:)?t>')) {
                    $value += Decode-XmlText $t.Groups[1].Value
                }
            } else {
                $v = Find-XmlMatches $cellXml '(?s)<(?:\w+:)?v(?:\s[^>]*)?>(.*?)</(?:\w+:)?v>' | Select-Object -First 1
                if ($v.Success) { $value = Decode-XmlText $v.Groups[1].Value }
                if ($type -eq 's' -and $value -match '^\d+$') {
                    $i = [int]$value
                    if ($i -ge $shared.Count) { throw "Ogiltigt shared-string-index $i i '$Path'." }
                    $value = $shared[$i]
                }
            }

            if ($type -in @('', 'n') -and
                $styleIndex -match '^\d+$') {

                $numFmtId = $styleFormats[[int]$styleIndex]

                if ($numFmtId -in @(14,15,16,17,18,19,20,21,22,45,47)) {
                    $value = Format-ExcelDate $value $numFmtId $date1904
                }
            }

            $cells[$index] = $value
        }
        $rows += ,$cells
    }
    if ($rows.Count -eq 0) { throw "Bladet '$actualSheet' är tomt."}

    $headers = @{}
    foreach ($index in $rows[0].Keys) {
        $header = [string]$rows[0][$index]
        if ($header -match '^\s*$') { continue }
        if ($headers.ContainsKey($header)) { throw "Duplicerad kolumnrubrik '$header' i '$Path'." }
        $headers[$header] = $index
    }
    if (-not $headers.ContainsKey($KeyColumn)) { trow "Nyckelkolumn '$KeyColumn' saknas i '$Path'." }
    
    $data = @{}
    $emptyKeys = 0
    $duplicates = 0

    for ($i = 1; $i -lt $rows.Count; $i++) {
        $raw = $rows[$i]
        $key = [string]$raw[$headers[$keyColumn]]
        if ($key -match '^\s*$') {
            $emptyKeys++
            $continue
        }
        if ($data.ContainsKey($key)) { 
            $duplicates++ 
            continue 
        }
        $record = @{}
        foreach ($header in $headers.Keys) { $record[$header] = [string]$raw[$headers[$header]] }
        $data[$key] = $record
    }
    if ($emptyKeys -gt 0 -or $duplicates -gt 0) {
        throw "'$Path': $emptykeys rader med tom nyckel och $duplicates duplicerade nycklar. Ingen jämföreslse utfördes."
    }
    return @{ Headers = $headers; Data = $data; Sheet = $actualSheet }
}

function Csv-Field {
    param([string]$Value)
    return '"' + $Value.Replace('"', '""') + '"'
}

if (-not (Get-Command tar.exe -ErrorAction SilentlyContinue)) {
    throw 'tar.exe saknas.'
}

$File1 = (Resolve-Path -LiteralPath $File1).Path
$File2 = (Resolve-Path -LiteralPath $File2).Path
if ($File1 -notmatch '(?i)\.xlsx$' -or $File2 -notmatch '(?i)\.xlsx$') {
    throw 'Endast .xlsx stöds.'
}

$a = Read-ExcelSheet $File1 $Sheet $KeyColumn
$b = Read-ExcelSheet $File2 $Sheet $KeyColumn

if ($Sheet -eq '' -and $a.Sheet -ne $b.Sheet){
    Write.Warning "Första bladet har olika namn: '$($a.Sheet)' respektive '$($b.Sheet)'."
}
$onlyA = @($a.Headers.Keys | Where-Object { $b.Headers.Keys -notcontains $_ })
$onlyB = @($b.Headers.Keys | Where-Object { $a.Headers.Keys -notcontains $_ })
if ($onlyA.Count -gt 0 -or $onlyB.Count -gt 0) {
    throw "Kolumnuppsättningarna skiljer sig. Endast fil 1: $($onlyA -join ', '). Endast fil 2: $($onlyB -join ', ')."
}

$lines = @('"Key","Status","Column","OldValue","NewValue"')
$added = 0; $removed = 0; $changed = 0; $unchanged = 0
foreach ($key in @((@($a.Data.Keys) + @($b.Data.Keys)) | Sort-Object -Unique)) {
    if (-not $a.Data.ContainsKey($key)) {
        $added++
        # $lines += ((Csv-Field $key) + ',"Added","","",""')
        foreach ($header in @($b.Headers.Keys | Sort-Object)) {
            if ($header -eq $KeyColumn) { continue }
            
            $newValue = [string]$b.Data[$key][$header]
            $lines += (
                (Csv-Field $key) +
                ',"Added",' +
                (Csv-Field $header) + ',' +
                (Csv-Field '') + ',' +
                (Csv-Field $newValue)
            )
        }
    } elseif (-not $b.Data.ContainsKey($key)) {
        $removed++
        # $lines += ((Csv-Field $key) + ',"Removed","","",""')

        foreach ($header in @($a.Headers.Keys | Sort-Object)) {
            if ($header -eq $KeyColumn) { continue }

            $oldValue = [string]$a.Data[$key][$header]

            $lines += (
                (Csv-Field $key) +
                ',"Removed",' +
                (Csv-Field $header) + ',' +
                (Csv-Field $oldValue) + ',' +
                (Csv-Field '')
            )
        }
    } else {
        $different = $false
        foreach ($header in @($a.Headers.Keys | Sort-Object)) {
            if ($header -eq $KeyColumn) { continue }
            $old = [string]$a.Data[$key][$header]
            $new = [string]$b.Data[$key][$header]
            if ($old -cne $new) {
                $different = $true
                $lines += ((Csv-Field $key) + ',"Changed",' + (Csv-Field $header) + ',' + (Csv-Field $old) + ',' + (Csv-Field $new))
            }
        }
        if ($different) { $changed++ } else { $unchanged++ }
    }
}

$lines | Set-Content -LiteralPath $Output -Encoding UTF8
Write-Host "Jämförelse klar. Blad: $($a.Sheet) / $($b.Sheet)"
Write-Host "Tillagda: $added | Borttagna: $removed | Ändrade: $changed | Oförändrade: $unchanged"
Write-Host "Resulatat: $Output"
