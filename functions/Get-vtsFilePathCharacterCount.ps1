function Get-vtsFilePathCharacterCount {
  <#
  .SYNOPSIS
      This script gets the character count of file names in a directory and exports the data to a CSV file.
  
  .DESCRIPTION
      The Get-vtsFilePathCharacterCount function takes a directory path and an output file path as parameters. It then retrieves all the items in the directory, counts the number of characters in each item's full name, and stores this information in an array. The array is then exported to a CSV file at the specified output file path.
  
  .PARAMETER directoryPath
      The path of the directory to get the file character count from. This parameter is mandatory.
  
  .PARAMETER outputFilePath
      The path of the CSV file to output the results to. This parameter is mandatory.
  
  .EXAMPLE
      Get-vtsFilePathCharacterCount -directoryPath "C:\path\to\directory" -outputFilePath "C:\path\to\outputfile.csv"
      This command gets the character count of all file names in the directory "C:\path\to\directory" and exports the data to the CSV file at "C:\path\to\outputfile.csv".
  
  .LINK
      File Management
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Please enter the directory path in the format 'C:\path\to\directory'")]
    [string]$directoryPath,
    [Parameter(Mandatory = $true, HelpMessage = "Please enter the output file path in the format 'C:\path\to\outputfile.csv'")]
    [string]$outputFilePath
  )

  Write-Host "Starting to get file character count from directory: $directoryPath"

  $items = Get-ChildItem -Path $directoryPath -Recurse

  Write-Host "Found $($items.Count) items in the directory"

  $output = @()

  foreach ($item in $items) {
    $output += New-Object PSObject -Property @{
      "Fullname"                         = $item.FullName
      "Number of characters in Fullname" = $item.FullName.Length
    }
  }

  Write-Host "Processed all items, now exporting to CSV file: $outputFilePath"

  $output | Export-Csv -Path $outputFilePath -NoTypeInformation

  Write-Host "Export completed successfully"
}

