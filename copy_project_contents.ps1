# PowerShell script to copy all project file contents to a single text file
# Usage: .\copy_project_contents.ps1

param(
    [string]$OutputFile = "project_contents.txt",
    [string]$ProjectPath = "."
)

# Get the absolute path of the project directory
$ProjectAbsolutePath = Resolve-Path $ProjectPath

# Initialize the output file
$OutputPath = Join-Path $ProjectAbsolutePath $OutputFile
"" | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host "Copying project contents to: $OutputPath"

# Function to get relative path
function Get-RelativePath {
    param(
        [string]$BasePath,
        [string]$FullPath
    )
    
    $RelativePath = $FullPath.Substring($BasePath.Length + 1)
    return $RelativePath.Replace('\', '/')
}

# Function to process files recursively
function Process-Files {
    param(
        [string]$Directory,
        [string]$BasePath
    )
    
    # Get all files in current directory (excluding the output file)
    $Files = Get-ChildItem -Path $Directory -File | Where-Object { $_.Name -ne $OutputFile }
    
    foreach ($File in $Files) {
        $RelativePath = Get-RelativePath -BasePath $BasePath -FullPath $File.FullName
        
        Write-Host "Processing: $RelativePath"
        
        # Write the separator and relative path
        "****$RelativePath****" | Out-File -FilePath $OutputPath -Append -Encoding UTF8
        
        try {
            # Read and write file content
            $Content = Get-Content -Path $File.FullName -Raw -Encoding UTF8
            if ($Content) {
                $Content | Out-File -FilePath $OutputPath -Append -Encoding UTF8 -NoNewline
            }
        }
        catch {
            "Error reading file: $_" | Out-File -FilePath $OutputPath -Append -Encoding UTF8
        }
        
        # Add extra newlines for separation
        "`n`n" | Out-File -FilePath $OutputPath -Append -Encoding UTF8 -NoNewline
    }
    
    # Process subdirectories recursively
    $Subdirectories = Get-ChildItem -Path $Directory -Directory
    foreach ($Subdirectory in $Subdirectories) {
        Process-Files -Directory $Subdirectory.FullName -BasePath $BasePath
    }
}

# Start processing from the project root
Process-Files -Directory $ProjectAbsolutePath -BasePath $ProjectAbsolutePath

Write-Host "Done! All file contents have been copied to: $OutputPath"
Write-Host "Total files processed: $((Get-Content $OutputPath | Select-String '^\*\*\*\*.*\*\*\*\*$').Count)"
