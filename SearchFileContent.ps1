param(
	[Parameter(Mandatory=$true)]
	[string]$SearchText
)

$path = (Get-Location).Path
$output="$env:TEMP\Officefiles.csv"

$textExt = @(
	".txt", ".csv", ".md", ".log",
	".xml", ".json", ".ps1"
)

$officeExt = @(
	".docx", ".xlsx", ".pptx"
)

$filesScanned = 0
$filesMatched = 0
$filesFailed=0

# Create result file
"FileName, FullPath, Extension, SearchText" |
    Set-content -LiteralPath $output -Encoding UTF8

function Test-OfficeContent {
    param(
        [string]$FilePath,
        [string]$Extension,
        [string]$SearchText
    )

    # List entries in the Office ZIP archive
    $entries = @(tar.exe -tf $FilePath 2>$null)
	
    if ($LASTEXITCODE -ne 0) {
        throw "Can not read the Office archive."
    }
	
	$found = $false

    foreach ($entry in $entries) {
        $include = $false

        switch ($Extension.ToLower()) {
            ".docx" {
                if (
					$entry -eq "word/document.xml" -or
                    $entry -match '^word/(header|footer|footnotes|endnotes)[0-9]*\.xml$'
				) {
                    $include = $true
                }
            }

            ".xlsx" {
                if (
					$entry -match '^xl/worksheets/sheet[0-9]+\.xml$' -or
                    $entry -eq "xl/sharedStrings.xml"
				) {
                    $include = $true
                }
            }

            ".pptx" {
                if (
					$entry -match '^ppt/slides/slide[0-9]+\.xml$'
				) {
                    $include = $true
                }
            }
        }

        if (-not $include) {
            continue
        }

        # Get the XML from the archive
        $xml = (tar.exe -xOf $FilePath $entry 2>$null) -join ""

        if ($LASTEXITCODE -ne 0) {
            throw "Cannt read XML entry"
        }

        # Preserve boundaries between paragraphs and cells
		$content = $xml -replace '</(w:p|a:p|row|c|si)>',  ''
		# Remove remaining XML tags without inserting spaces
        $content = $content -replace '<[^>]+>',  ''

        # Decode common XML entities
        $content = $content.Replace('&amp;', '&')
        $content = $content.Replace('&lt;', '<')
        $content = $content.Replace('&gt;', '>')
        $content = $content.Replace('&quot;', '"')
        $content = $content.Replace('&apos;', "'")

        # Search without regard to uppercase or lowercase letters 
        if (
			$content | 
			Select-String -SimpleMatch -Pattern $SearchText -Quiet
		) {
            $found = $true
			break
        }
    }

    return $found
}

# Scan files recursively
Get-ChildItem -LiteralPath $path -Recurse -File `
	-ErrorAction continue | ForEach-Object {
		
	$file = $_

    # Avoid searchin the own result file
    if ($file.FullName -eq $output) {
        return
    }

    if ($file.Extension -notin ($textExt + $officeExt)) {
        return
    }

	$filesScanned++
    $found = $false
	
    try {
	    if ($file.Extension -in $textExt) {
			
		    $found = Select-String `
                -LiteralPath $file.FullName `
                -SimpleMatch `
                -Pattern $SearchText `
                -Quiet `
                -ErrorAction Continue 
        }
	    else {
            $found = Test-OfficeContent `
                -FilePath $file.FullName `
                -Extension $file.Extension `
                -SearchText $SearchText
	    }
    }
    catch {
        $filesFailed++
        Write-Warning "Can not read: $($file.FullName)"
		Write-Warning "Error:$($_.Exception.Message)"
    }

    if ($found) {
        $filesMatched++

        $name = $file.Name.Replace('"', '""')
        $full = $file.FullName.Replace('"', '""')
        $term = $SearchText.Replace('"', '""')

        '"{0}","{1}","{2}","{3}"' -f `
            $name, $full, $file.Extension, $term |
            Add-Content -Path $output -Encoding UTF8
            
        Write-Host "HIT: $($file.FullName)"
    }
         
}

Write-Host ""
Write-Host "======================================================"
Write-Host " SEARCH COMPLETED"
Write-Host "======================================================"
Write-Host "Search path: $path"
Write-Host "Search string: $SearchText"
Write-host "Files scanned: $filesScanned"
Write-Host "Files matched: $filesMatched"
Write-Host "Files failed: $filesFailed"
Write-Host "CSV-file: $output"
Write-Host "======================================================"

