function Remove-FTA {
  <#
  .DESCRIPTION
      Remove the file type association for a given application.
  
  .NOTES
      Version    : 1.2.0
      Author(s)  : Danyfirex & Dany3j
      Credits    : https://bbs.pediy.com/thread-213954.htm
                   LMongrain - Hash Algorithm PureBasic Version
      License    : MIT License
      Copyright  : 2022 Danysys. <danysys.com>
  
  .LINK
  File Association Management
      #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [Alias("ProgId")]
    [String]
    $ProgramPath,

    [Parameter(Mandatory = $true)]
    [String]
    $Extension
  )
  
  function local:Remove-UserChoiceKey {
    param (
      [Parameter( Position = 0, Mandatory = $True )]
      [String]
      $Key
    )

    $code = @'
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Win32;
    
    namespace Registry {
      public class Utils {
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern int RegOpenKeyEx(UIntPtr hKey, string subKey, int ulOptions, int samDesired, out UIntPtr hkResult);
    
        [DllImport("advapi32.dll", SetLastError=true, CharSet = CharSet.Unicode)]
        private static extern uint RegDeleteKey(UIntPtr hKey, string subKey);

        public static void DeleteKey(string key) {
          UIntPtr hKey = UIntPtr.Zero;
          RegOpenKeyEx((UIntPtr)0x80000001u, key, 0, 0x20019, out hKey);
          RegDeleteKey((UIntPtr)0x80000001u, key);
        }
      }
    }
'@

    try {
      Add-Type -TypeDefinition $code
    }
    catch {}

    try {
      [Registry.Utils]::DeleteKey($Key)
    }
    catch {} 
  } 

  function local:Update-Registry {
    $code = @'
    [System.Runtime.InteropServices.DllImport("Shell32.dll")] 
    private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
    public static void Refresh() {
        SHChangeNotify(0x8000000, 0, IntPtr.Zero, IntPtr.Zero);    
    }
'@ 

    try {
      Add-Type -MemberDefinition $code -Namespace SHChange -Name Notify
    }
    catch {}

    try {
      [SHChange.Notify]::Refresh()
    }
    catch {} 
  }

  if (Test-Path -Path $ProgramPath) {
    $ProgId = "SFTA." + [System.IO.Path]::GetFileNameWithoutExtension($ProgramPath).replace(" ", "") + $Extension
  }
  else {
    $ProgId = $ProgramPath
  }

  try {
    $keyPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$Extension\UserChoice"
    Write-Verbose "Remove User UserChoice Key If Exist: $keyPath"
    Remove-UserChoiceKey $keyPath

    $keyPath = "HKCU:\SOFTWARE\Classes\$ProgId"
    Write-Verbose "Remove Key If Exist: $keyPath"
    Remove-Item -Path $keyPath -Recurse -ErrorAction Stop | Out-Null
    
  }
  catch {
    Write-Verbose "Key No Exist: $keyPath"
  }

  try {
    $keyPath = "HKCU:\SOFTWARE\Classes\$Extension\OpenWithProgids"
    Write-Verbose "Remove Property If Exist: $keyPath Property $ProgId"
    Remove-ItemProperty -Path $keyPath -Name $ProgId -ErrorAction Stop | Out-Null
    
  }
  catch {
    Write-Verbose "Property No Exist: $keyPath Property: $ProgId"
  } 

  Update-Registry
  Write-Output "Removed: $ProgId" 
}

