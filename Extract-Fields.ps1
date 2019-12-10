param (
    [string]$PbiPath = ".\",
    [string]$OutputCsvPath = "fields.csv",
    [switch]$Force = $false
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

# From https://stackoverflow.com/a/34559554
function New-TemporaryDirectory {
    $parent = [System.IO.Path]::GetTempPath()
    [string] $name = [System.Guid]::NewGuid()
    New-Item -ItemType Directory -Path (Join-Path $parent $name)
}

# Check if output file exists and abort if file exists
if ((Test-Path $OutputCsvPath -PathType Leaf) -and ($Force -eq $false)) {
    Write-Warning "$($OutputCsvPath) already exists. Use -Force if to overwrite"
    exit
}

# Get all pbix and pbit files
$files = Get-ChildItem -Path $PbiPath -Include *.pbi[xt] -Recurse

$allFields = foreach ($file in $files) {
    # Create temp directory where we will copy the file for unzipping
    $tempDir = New-TemporaryDirectory

    # Copy the file to the temp directory
    Copy-Item $file -Destination $tempDir

    # Location of file we just copied
    $oldName = Join-Path $tempDir $file.Name

    # Hard-coded file.zip as it doesn't matter what the filename is called
    $newName = Join-Path $tempDir "file.zip"

    # Replace .pbi[xt] with .zip
    # This is because Expand-Archive only works on .zip files
    Rename-Item -Path $oldName -NewName $newName

    # Extract file.zip
    Expand-Archive $newName -DestinationPath $tempDir

    # Layout is where references to fields are
    $layoutFile = Join-Path $tempDir "Report\Layout"

    # Convert layout to JSON object
    $json = Get-Content -Raw $layoutFile -Encoding Unicode | Out-String | ConvertFrom-Json

    # Extract fields and tables from the filters section
    $fields = foreach ($filter in $json.sections.filters) {
        $expressions = $filter | Out-String | ConvertFrom-Json
        
        foreach ($expression in $expressions.expression) {
            $table = $expression.Column.Expression.SourceRef.Entity
            $field = $expression.Column.Property
            New-Object psobject -Property @{
                File = $file
                Table = $table
                Field = $field
            }
        }
    }

    # Extract fields and tables from the visual queries section
    $fields += foreach ($visualQuery in $json.sections.visualContainers.query) {
        $query = $visualQuery | Out-String | ConvertFrom-Json
        $name = $query.Commands.SemanticQueryDataShapeCommand.Query.Select.Name
        $names = $query.Commands.SemanticQueryDataShapeCommand.Query.Select.Name | Where-Object {$_ -ne $null} | Where-Object {$_.Contains(".")}

        foreach ($name in $names) {
            # If field is aggregated, e.g. Min(table.field), then strip out aggregating function
            $tableStartIndex = $name.LastIndexOf("(")
            if ($tableStartIndex -gt -1) {
                $tableEndIndex = $name.IndexOf(")")
                $name = $name.Substring($tableStartIndex + 1, $tableEndIndex - $tableStartIndex - 1)
            }

            # Split table.field
            $split = $name.Split(".")
            
            New-Object psobject -Property @{
                File = $file
                Table = $split[0]
                Field = $split[1]
            }
        }
    }

    # As a catch all, find all fields (without corresponding table)
    $matches = Select-String -Path $layoutFile -Encoding Unicode -Pattern '\\"Property\\":\\"(.*?)\\"' -AllMatches

    # Select the field name from the capture group
    $matches = $matches.Matches.Groups.Groups | Where-Object Name -eq 1 | Select-Object -Property Value -Unique | Sort-Object Value

    # Exclude fields where we have previously matched to a table
    # This assumes field names are fairly unique in the dataset
    $matches = $matches | Where-Object { ($fields | foreach {$_.Field}) -notcontains $_.Value }

    $fields += foreach($match in $matches) {
        New-Object psobject -Property @{
            File = $file
            Table = ""
            Field = $match.Value
        }
    }

    # Remove the temporary directory
    Remove-Item $tempDir -Recurse

    # Return the fields to the outer scope to be collected by $allFIelds
    $fields
}

# Export the data to CSV
$allFields | Sort-Object File, Table, Field -Unique | Export-Csv -Path $OutputCsvPath -NoTypeInformation
