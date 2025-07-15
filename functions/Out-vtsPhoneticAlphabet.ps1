function Out-vtsPhoneticAlphabet {
  <#
  .Description
  Converts strings to the phonetic alphabet.
  .EXAMPLE
  PS> "RandomString" | Out-vtsPhoneticAlphabet
  
  Output:
  ROMEO
  alfa
  november
  delta
  oscar
  mike
  SIERRA
  tango
  romeo
  india
  november
  golf
  
  .LINK
  Utilities
  #>
  [CmdletBinding()]
  [OutputType([String])]
  Param
  (
    # Input string to convert
    [Parameter(Mandatory = $true, 
      ValueFromPipeline = $true,
      Position = 0)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $InputObject
  )
  $result = @()
  $nato = @{
    '0' = '(ZERO)'
    '1' = '(ONE)'
    '2' = '(TWO)'
    '3' = '(THREE)'
    '4' = '(FOUR)'
    '5' = '(FIVE)'
    '6' = '(SIX)'
    '7' = '(SEVEN)'
    '8' = '(EIGHT)'
    '9' = '(NINE)'
    'a' = 'alfa'
    'b' = 'bravo'
    'c' = 'charlie'
    'd' = 'delta'
    'e' = 'echo'
    'f' = 'foxtrot'
    'g' = 'golf'
    'h' = 'hotel'
    'i' = 'india'
    'j' = 'juliett'
    'k' = 'kilo'
    'l' = 'lima'
    'm' = 'mike'
    'n' = 'november'
    'o' = 'oscar'
    'p' = 'papa'
    'q' = 'quebec'
    'r' = 'romeo'
    's' = 'sierra'
    't' = 'tango'
    'u' = 'uniform'
    'v' = 'victor'
    'w' = 'whiskey'
    'x' = 'xray'
    'y' = 'yankee'
    'z' = 'zulu'
    '.' = '(PERIOD)'
    '-' = '(DASH)'
  }

  $chars = ($InputObject).ToCharArray()

  foreach ($char in $chars) {
    switch -Regex -CaseSensitive ($char) {
      '\d' {
        $result += ($nato["$char"])
        break
      }
      '[a-z]' {
        $result += ($nato["$char"]).ToLower()
        break
      }
      '[A-Z]' {
        $result += ($nato["$char"]).ToUpper()
        break
      }

      Default { $result += $char }
    }
  }
  $result
}

