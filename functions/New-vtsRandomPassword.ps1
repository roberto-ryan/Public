function New-vtsRandomPassword {
  <#
  .Description
  Generates a random 12 character password and copies it to the clipboard.
  If the -Easy switch is specified, generates a password using two random words, a random number and a random symbol.
  Now, it also adds a random preposition between the two words for the Easy password.
  .EXAMPLE
  PS> New-vtsRandomPassword
  
  Output:
  Random Password Copied to Clipboard
  
  .EXAMPLE
  PS> New-vtsRandomPassword -Easy -WordListPath "C:\path\to\wordlist.csv"
  
  Output:
  Easy Random Password Copied to Clipboard
  
  .LINK
  Utilities
  #>
  param(
    [switch]$Easy,
    [string]$WordListPath = "$env:temp\wordlist.csv"
  )

  $numbers = 0..9
  $symbols = '!', '@', '#', '$', '%', '*', '?', '+', '='
  $prepositions = 'on', 'in', 'at', 'by', 'up', 'to', 'of', 'off', 'for', 'out', 'via'
  $number = $numbers | Get-Random
  $symbol = $symbols | Get-Random

  if ($Easy) {
    if (!(Test-Path $WordListPath)) {
      Invoke-WebRequest -uri "https://raw.githubusercontent.com/roberto-ryan/Public/main/wordlist.csv" -UseBasicParsing -OutFile $WordListPath
    }
    $words = Import-Csv -Path $WordListPath | ForEach-Object { $_.Word } 
    $randomPreposition = ($prepositions | Get-Random).ToUpper()
    $randomWord1 = $words | Get-Random
    $randomWord2 = $words | Get-Random
    $NewPW = $randomWord1 + $randomPreposition + $randomWord2 + $number + $symbol
  }
  else {
    $string = ( -join ((0x30..0x39) + ( 0x41..0x5A) + ( 0x61..0x7A) |
        Get-Random -Count 12  |
        ForEach-Object { [char]$_ }))
    $number = $numbers | Get-Random
    $symbol = $symbols | Get-Random
    $NewPW = $string + $number + $symbol
  }

  $NewPW | Set-Clipboard

  if ($Easy) {
    Write-Output "Easy Random Password Copied to Clipboard - $NewPW"
  }
  else {
    Write-Output "Random Password Copied to Clipboard - $NewPW"
  }
}

