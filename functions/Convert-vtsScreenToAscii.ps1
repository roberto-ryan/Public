function Convert-vtsScreenToAscii {
  <#
  .SYNOPSIS
  This script takes a screenshot of the current display and converts it to ASCII art to be displayed directly in the console.
  
  .DESCRIPTION
  The script uses the Get-vtsScreenshot function to capture a screenshot and save it to a specified directory. It then uses the ImageToAscii class to convert the screenshot into ASCII art. The ASCII art is then printed to the console.
  
  .PARAMETER ImageDirectory
  The directory where the screenshot will be saved. Default is "C:\temp\".
  
  .PARAMETER NumOfImages
  The number of screenshots to take and convert to ASCII art. Default is 3.
  
  .PARAMETER SleepInterval
  The interval in seconds between each screenshot. Default is 1.
  
  .EXAMPLE
  PS C:\> .\screen-to-ascii.ps1 -ImageDirectory "C:\screenshots\" -NumOfImages 5 -SleepInterval 2
  This example will take 5 screenshots at an interval of 2 seconds each, save them to the "C:\screenshots\" directory, and convert each one to ASCII art.
  
  .NOTES
  The ASCII art is generated with a fixed width of 100 characters. This is to maintain the aspect ratio of the ASCII characters.
  
  .LINK
  Utilities
  
  #>
  param (
    [string]$ImageDirectory = "$env:temp\",
    [int]$NumOfImages = 100,
    [int]$SleepInterval = 0
  )

  if ((whoami) -eq "nt authority\system") {
    Write-Error "Must run script as logged in user. Running as system doesn't work."
  }
  else {
    # Define the ImageToAscii class
    Add-Type -TypeDefinition @"
using System;
using System.Drawing;
public class ImageToAscii {
  public static string ConvertImageToAscii(string imagePath, int width) {
      Bitmap image = new Bitmap(imagePath, true);
      image = GetResizedImage(image, width);
      return ConvertToAscii(image);
  }
  private static Bitmap GetResizedImage(Bitmap original, int width) {
      int height = (int)(original.Height * ((double)width / original.Width) / 2); // Adjust for aspect ratio of ASCII characters
      var resized = new Bitmap(original, new Size(width, height));
      return resized;
  }
  private static string ConvertToAscii(Bitmap image) {
      string ascii = "";
      for (int h = 0; h < image.Height; h++) {
          for (int w = 0; w < image.Width; w++) {
              Color pixelColor = image.GetPixel(w, h);
              int grayScale = (pixelColor.R + pixelColor.G + pixelColor.B) / 3;
              ascii += GetAsciiCharForGrayscale(grayScale);
          }
          ascii += "\\n";
      }
      return ascii;
  }
  private static char GetAsciiCharForGrayscale(int grayScale) {
      string asciiChars = "@%#*+=-:. ";
      return asciiChars[grayScale * asciiChars.Length / 256];
  }
}
"@ -ReferencedAssemblies System.Drawing *>$null -ErrorAction SilentlyContinue

    try {
      if (!(Test-Path -Path $ImageDirectory)) {
        New-Item -ItemType Directory -Force -Path $ImageDirectory
      }
    
    
      for ($i = 1; $i -le $NumOfImages; $i++) {
        $imagePath = Join-Path -Path $ImageDirectory -ChildPath "image_$i.png"
        Get-vtsScreenshot -Path $imagePath *>$null
    
          
        # Convert the image to ASCII art
        $asciiArt = [ImageToAscii]::ConvertImageToAscii($imagePath, 100) # Adjust the width for ASCII character aspect ratio
          
        # Print the ASCII art to the host a chunk at a time
        $lines = $asciiArt -split "\\n"
          
        foreach ($line in $lines) {
          if ($line.Length -gt 100) {
            # If the line is longer than 100 characters, split it into chunks
            $chunks = $line -split "(.{100})", -1, 'RegexMatch'
            Clear-Host
            foreach ($chunk in $chunks) {
              if ($chunk -ne "") {
                Write-Host $chunk
              }
            }
          }
          else {
            # If the line is not longer than 100 characters, just print it
            Write-Host $line
          }
    
        }
        Start-Sleep $SleepInterval
      }
    
    }
    finally {
      <#Do this after the try block regardless of whether an exception occurred or not#>
      Remove-Item "$ImageDirectory\image_*.png" -Recurse -Force -Confirm:$false
    }

  }

}

