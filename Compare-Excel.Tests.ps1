$scriptPath = Join-Path $PSScriptRoot 'Compare-Excel.ps1'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function New-TestWorkbook {
    param([string]$Path, [string]$Rows, [string]$SheetName = 'Sheet1')

    $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $entries = @{
            'xl/workbook.xml' = "<workbook><sheets><sheet name=`"$SheetName`" r:id=`"rId1`"/></sheets></workbook>"
            'xl/_rels/workbook.xml.rels' = '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>'
            'xl/worksheets/sheet1.xml' = "<worksheet><sheetData>$Rows</sheetData></worksheet>"
        }
        foreach ($name in $entries.Keys) {
            $writer = New-Object System.IO.StreamWriter($archive.CreateEntry($name).Open())
            try { $writer.Write($entries[$name]) } finally { $writer.Dispose() }
        }
    } finally { $archive.Dispose() }
}

Describe 'Compare-Excel' {
    It 'compares physical rows including gaps, ID values and reordered columns' {
        $old = Join-Path $TestDrive 'old.xlsx'
        $new = Join-Path $TestDrive 'new.xlsx'
        $output = Join-Path $TestDrive 'result.csv'
        New-TestWorkbook $old @'
<row r="1"><c r="A1" t="inlineStr"><is><t>Id</t></is></c><c r="B1" t="inlineStr"><is><t>Name</t></is></c></row>
<row r="2"><c r="B2" t="inlineStr"><is><t>Alice</t></is></c></row>
<row r="4"><c r="A4" t="inlineStr"><is><t>same</t></is></c><c r="B4" t="inlineStr"><is><t>Carol</t></is></c></row>
'@
        New-TestWorkbook $new @'
<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Id</t></is></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>Alicia</t></is></c><c r="B2" t="inlineStr"><is><t>same</t></is></c></row>
<row r="3"><c r="A3" t="inlineStr"><is><t>Bob</t></is></c></row>
<row r="4"><c r="A4" t="inlineStr"><is><t>Carol</t></is></c><c r="B4" t="inlineStr"><is><t>same</t></is></c></row>
<row r="5"><c r="A5" t="inlineStr"><is><t>David</t></is></c></row>
'@
        & $scriptPath -File1 $old -File2 $new -Output $output
        $result = @(Import-Csv $output)
        $result.Count | Should Be 5
        @($result | Where-Object { $_.Row -eq '2' -and $_.Status -eq 'Changed' -and $_.Column -eq 'Id' -and $_.NewValue -eq 'same' }).Count | Should Be 1
        @($result | Where-Object { $_.Row -eq '2' -and $_.Column -eq 'Name' -and $_.OldValue -eq 'Alice' -and $_.NewValue -eq 'Alicia' }).Count | Should Be 1
        @($result | Where-Object { $_.Row -eq '3' -and $_.Status -eq 'Changed' -and $_.Column -eq 'Name' -and $_.OldValue -eq '' -and $_.NewValue -eq 'Bob' }).Count | Should Be 1
        @($result | Where-Object { $_.Row -eq '4' }).Count | Should Be 0
        @($result | Where-Object { $_.Row -eq '5' -and $_.Status -eq 'Added' }).Count | Should Be 2
    }

    It 'reports rows past the end of the new sheet as removed' {
        $old = Join-Path $TestDrive 'long.xlsx'
        $new = Join-Path $TestDrive 'short.xlsx'
        $output = Join-Path $TestDrive 'removed.csv'
        New-TestWorkbook $old '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Alice</t></is></c></row>'
        New-TestWorkbook $new '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c></row>'
        & $scriptPath -File1 $old -File2 $new -Output $output
        $result = @(Import-Csv $output)
        $result.Count | Should Be 1
        $result[0].Row | Should Be '3'
        $result[0].Status | Should Be 'Removed'
        $result[0].OldValue | Should Be 'Alice'
    }

    It 'still compares by key when a key column is supplied' {
        $old = Join-Path $TestDrive 'key-old.xlsx'
        $new = Join-Path $TestDrive 'key-new.xlsx'
        $output = Join-Path $TestDrive 'key.csv'
        New-TestWorkbook $old '<row r="1"><c r="A1" t="inlineStr"><is><t>Id</t></is></c><c r="B1" t="inlineStr"><is><t>Name</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>one</t></is></c><c r="B2" t="inlineStr"><is><t>Alice</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>two</t></is></c><c r="B3" t="inlineStr"><is><t>Bob</t></is></c></row>'
        New-TestWorkbook $new '<row r="1"><c r="A1" t="inlineStr"><is><t>Id</t></is></c><c r="B1" t="inlineStr"><is><t>Name</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>two</t></is></c><c r="B2" t="inlineStr"><is><t>Bobby</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>one</t></is></c><c r="B3" t="inlineStr"><is><t>Alice</t></is></c></row>'
        & $scriptPath -File1 $old -File2 $new -KeyColumn Id -Output $output
        $result = @(Import-Csv $output)
        $result.Count | Should Be 1
        $result[0].Key | Should Be 'two'
        $result[0].Status | Should Be 'Changed'
        $result[0].OldValue | Should Be 'Bob'
        $result[0].NewValue | Should Be 'Bobby'
    }

    It 'rejects duplicate keys when a key column is supplied' {
        $path = Join-Path $TestDrive 'duplicates.xlsx'
        $output = Join-Path $TestDrive 'duplicates.csv'
        New-TestWorkbook $path '<row r="1"><c r="A1" t="inlineStr"><is><t>Id</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>same</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>same</t></is></c></row>'
        { & $scriptPath -File1 $path -File2 $path -KeyColumn Id -Output $output } | Should Throw 'duplicerade nycklar'
    }
}
