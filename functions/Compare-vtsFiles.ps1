function Compare-vtsFiles {
  <#
  .SYNOPSIS
  This script compares the files in two directories and generates a report of their SHA1 hashes.
  
  .DESCRIPTION
  The Compare-vtsFiles function takes in two parameters, the source folder and the destination folder. It calculates the SHA1 hash of each file in both folders and compares them. If the hashes match, it means the files are identical. If they don't, the files are different. The function generates a report of the comparison results, including the file paths and their hashes. The report is saved in a CSV file in the TEMP directory.
  
  .PARAMETER SourceFolder
  The path of the source folder.
  
  .PARAMETER DestinationFolder
  The path of the destination folder.
  
  .PARAMETER ReportPath
  The path where the report will be saved. By default, it is saved in the TEMP directory.
  
  .EXAMPLE
  Compare-vtsFiles -SourceFolder "C:\Source" -DestinationFolder "C:\Destination"
  
  This will compare the files in the Source and Destination folders and generate a report in the TEMP directory.
  
  .EXAMPLE
  Compare-vtsFiles -SourceFolder "C:\Source" -DestinationFolder "C:\Destination" -ReportPath "C:\Reports\Hashes.csv"
  
  This will compare the files in the Source and Destination folders and generate a report in the specified path.
  
  .LINK
  File Management
  #>
  param (
    [string]$SourceFolder,
    [string]$DestinationFolder,
    [string]$ReportPath = "$env:TEMP\VTS\$(Get-Date -f yyyy-MM-dd-hhmmss)_Hashes.csv"
  )

  # Initialize a new list to store the results
  $result = [System.Collections.Generic.List[object]]::new()

  # Define a script block to process each file
  $sb = {
    process {
      # Ignore 'Thumbs.db' files
      if ($_.Name -eq 'Thumbs.db') { return }

      # Create a custom object with file properties
      [PSCustomObject]@{
        h  = (Get-FileHash $_.FullName -Algorithm SHA1).Hash # File hash
        n  = $_.Name # File name
        fn = $_.fullname # Full file path
      }
    }
  }

  # Get all files from the source and destination folders
  $sourceFiles = Get-ChildItem $SourceFolder -Recurse -File | & $sb
  $destinationFiles = Get-ChildItem $DestinationFolder -Recurse -File | & $sb

  # Process each file in the source folder
  foreach ($file in $sourceFiles) {
    # If the file exists in the destination folder
    if ($destinationFile = $destinationFiles | Where-Object { $_.n -eq $file.n }) {
      # Create a custom object with source and target file properties
      $comparisonResult = [PSCustomObject]@{
        SourceFilePath = $file.fn
        SourceFileHash = $file.h
        TargetFilePath = $destinationFile.fn
        TargetFileHash = $destinationFile.h
        Status         = if ($file.h -eq $destinationFile.h) { 'Hashes Match' } else { 'Hashes Do Not Match' }
      }
    }
    else {
      # If the file does not exist in the destination folder
      $comparisonResult = [PSCustomObject]@{
        SourceFilePath = $file.fn
        SourceFileHash = $file.h
        TargetFilePath = $null
        TargetFileHash = $null
        Status         = 'File not found in destination'
      }
    }

    # Add the comparison result to the result list
    $result.Add($comparisonResult)
  }

  $ReportDirectory = Split-Path -Path $ReportPath -Parent
  if (!(Test-Path -Path $ReportDirectory)) {
    New-Item -ItemType Directory -Path $ReportDirectory -Force | Out-Null
  }

  # Output the result list in a table format
  $result | Export-Csv -NoTypeInformation -Path $ReportPath -Force

  Write-Host "Report exported to $ReportPath"
}

