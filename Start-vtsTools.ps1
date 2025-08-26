[CmdletBinding()]
param(
  [string]$FunctionsPath = $PSScriptRoot,
  # Theme variant and Unicode controls
  [ValidateSet('rose-pine', 'rose-pine-moon', 'rose-pine-dawn', 'rose-pine-contrast', 'dracula', 'nord', 'one-dark', 'tokyo-night', 'gruvbox', 'catppuccin-mocha', 'solarized-dark')]
  [string]$ThemeVariant = 'rose-pine',
  [ValidateSet('minimal', 'classic')]
  [string]$ThemeStyle = 'classic',
  [switch]$NoUnicode,
  # If provided, this script can download the repo ZIP, extract it, and run Start-vtsTools.ps1 from there.
  [string]$BootstrapFromRepoUrl,
  # Preferred branch to download when bootstrapping; may be overridden by URL if it contains a branch.
  [string]$Branch = 'main',
  # Relative path to the start script inside the repo when bootstrapping.
  [string]$StartScriptRelPath = 'Start-vtsTools.ps1',
  # Base folder for extraction; default chosen automatically if not provided.
  [string]$InstallBase,
  # Force re-download by deleting any existing extracted folder first.
  [switch]$Reinstall,
  # Build model and print a summary, skip launching the interactive TUI.
  [switch]$ValidateOnly,
  # Populate the UI without importing or dot-sourcing files (prevents side-effects)
  [bool]$ScanSafe = $false,
  # How to format category labels for URL/GitHub-derived sources: Full = owner/repo@branch, OwnerRepo = owner/repo, Repo = repo only
  [ValidateSet('Full', 'OwnerRepo', 'Repo')]
  [string]$CategoryLabelStyle = 'OwnerRepo',
  # Use repository folders to build categories first (overrides .LINK/.help when present)
  [bool]$PreferFolderCategory = $true,
  # When bootstrapping from a repo URL, scan and show the embedded TUI instead of executing the repo's start script
  [bool]$BootstrapScanOnly = $true,
  # How many folder segments to include as a category (after skipping ignored prefixes)
  [ValidateRange(1, 5)]
  [int]$CategoryDepth = 2,
  # Folder names to skip at the beginning of relative paths when deriving categories
  [string[]]$CategoryIgnoreFolders = @('functions', 'function', 'scripts', 'script', 'src', 'source', 'powershell', 'pwsh', 'ps', 'bin', 'build', '.github', '.vscode', 'lib', 'modules', 'module', 'samples', 'examples', 'test', 'tests', 'docs', 'documentation'),
  # Optional map to rename specific folder names to nicer labels (case-insensitive keys)
  [hashtable]$CategoryRenameMap = @{},
  # Optional map to override final categories by exact or regex match
  # Keys starting with 'Regex:' are treated as regex patterns (case-insensitive example shown).
  # These defaults are conservative and can be overridden by passing your own map.
  [hashtable]$CategoryOverrideMap = [ordered]@{
    # Microsoft 365
    'Regex:(?i)\b(M365|O365|Office365|ExchangeOnline|SharePointOnline|OneDrive|Teams|Outlook|Planner|PowerPlatform|PowerApps|PowerAutomate)\b'                  = 'Microsoft 365'

    # Identity
    'Regex:(?i)\b(ActiveDirectory|ADFS|ADDS|DomainController|SAMAccount|Kerberos|LDAP|AAD|AzureAD|EntraID)\b'                                                   = 'Identity'

    # Security
    'Regex:(?i)\b(Security|SecPol|ACL|Permissions|Credential|Secrets|PKI|Certificate|TLS|SSL|Firewall|Malware|Virus|Defender|AppLocker|BitLocker|Encryption)\b' = 'Security'

    # Endpoint Management
    'Regex:(?i)\b(Intune|SCCM|ConfigMgr|EndpointManager|MDM|GPO|GroupPolicy)\b'                                                                                 = 'Endpoint Mgmt'

    # Azure
    'Regex:(?i)\b(Azure|ARM|ResourceGroup|VMSS|AKS|AppService|KeyVault|StorageAccount|CosmosDB|LogicApp|FunctionApp|VNet|NSG)\b'                                = 'Azure'

    # Dev / Work Mgmt
    'Regex:(?i)\b(Jira|Confluence|DevOps|Agile|Scrum|Kanban|Trello|Asana)\b'                                                                                    = 'Dev/Work Mgmt'

    # Virtualization
    'Regex:(?i)\b(Hyper-V|VMware|vSphere|ESXi|VirtualBox|VHDX?|Snapshot|Checkpoint)\b'                                                                          = 'Virtualization'

    # Containers
    'Regex:(?i)\b(Docker|Podman|Containerd|K8s|Kubernetes|Helm|Image|Container)\b'                                                                              = 'Containers'

    # DevOps / IaC
    'Regex:(?i)\b(DevOps|Terraform|Ansible|Chef|Puppet|Bicep|ARMTemplate|CI/CD|Pipeline|Jenkins|Octopus)\b'                                                     = 'DevOps/IaC'

    # Databases
    'Regex:(?i)\b(SQLServer|MSSQL|Postgres|PostgreSQL|MySQL|MariaDB|OracleDB|MongoDB|Redis|Database|SQLite)\b'                                                  = 'Databases'

    # Auth Standards
    'Regex:(?i)\b(OAuth|OIDC|SAML|JWT|FIDO2|MFA|2FA|SSO|AuthN|AuthZ)\b'                                                                                         = 'Auth Standards'

    # Cloud Vendors
    'Regex:(?i)\b(AWS|AmazonWebServices|EC2|S3|GCP|GoogleCloud|BigQuery|CloudRun)\b'                                                                            = 'Cloud Vendors'

    # Networking
    'Regex:(?i)\b(Network|NetCfg|DNS|DHCP|IPConfig|Ping|Traceroute|Subnet|Routing|Switch|Router|WiFi|NAT|Port|TCP|UDP|SSLVPN|VPN)\b'                            = 'Networking'

    # Languages
    'Regex:(?i)\b(PowerShell|Bash|Python|CSharp|C#|JavaScript|TypeScript|GoLang?|Rust|Perl|Ruby)\b'                                                             = 'Languages'

    # Backup and DR
    'Regex:(?i)\b(Backup|Restore|Recovery|Snapshot|Replication|Failover|DisasterRecovery|DR|Veeam)\b'                                                           = 'Backup and DR'

    # Email
    'Regex:(?i)\b(Exchange|SMTP|IMAP|POP3|Mailbox|MailFlow|SendMail)\b'                                                                                         = 'Email'

    # --- Granular Sub-Categories ---

    # Help/Discovery
    'Regex:(?i)\b(Help|Get-Help|About_|Info|Discover|WhatIf)\b'                                                                                                 = 'Help/Discovery'

    # Processes
    'Regex:(?i)\b(Process|Tasklist|Taskkill|Get-Process|ProcMon|Handle)\b'                                                                                      = 'Processes'

    # Services
    'Regex:(?i)\b(Service|Get-Service|Set-Service|Start-Service|Stop-Service|Restart-Service)\b'                                                                = 'Services'

    # Files/Directories
    'Regex:(?i)\b(File|Folder|Directory|Path|Copy-Item|Move-Item|Remove-Item|Rename-Item|New-Item|Get-ChildItem|Tree)\b'                                        = 'Files/Directories'

    # Registry
    'Regex:(?i)\b(Registry|RegKey|HKLM|HKCU|HKCR|HKU|HKCC|Get-ItemProperty|Set-ItemProperty)\b'                                                                 = 'Registry'

    # Events/Logs
    'Regex:(?i)\b(EventLog|Get-EventLog|Get-WinEvent|LogName|ApplicationLog|SystemLog|Audit)\b'                                                                 = 'Events/Logs'

    # Hardware/Devices
    'Regex:(?i)\b(Device|PnP|Driver|Hardware|Disk|Volume|Partition|USB|PrinterPort|Monitor|Battery|Adapter)\b'                                                  = 'Hardware/Devices'

    # Printing
    'Regex:(?i)\b(Print|Printer|PrintJob|Spooler|PrintQueue)\b'                                                                                                 = 'Printing'

    # Updates
    'Regex:(?i)\b(Update|Patch|WUInstall|WindowsUpdate|KB\d+)\b'                                                                                                = 'Updates'

    # Users/Groups
    'Regex:(?i)\b(User|Group|LocalUser|LocalGroup|Account|SID|Profile|Credential|NTUser)\b'                                                                     = 'Users/Groups'

    # Performance
    'Regex:(?i)\b(PerfMon|Performance|Counter|ResourceMonitor|CPU|Memory|DiskIO|Latency|Benchmark)\b'                                                           = 'Performance'

    # Power
    'Regex:(?i)\b(PowerPlan|Battery|Sleep|Hibernate|Shutdown|Restart|Reboot|UPS)\b'                                                                             = 'Power'

    # Display
    'Regex:(?i)\b(Display|Resolution|Monitor|Screen|Graphics|DPI|Color|Brightness)\b'                                                                           = 'Display'

    # Input
    'Regex:(?i)\b(Keyboard|Mouse|Input|HID|Touchpad|Tablet|Pen)\b'                                                                                              = 'Input'

    # Audio
    'Regex:(?i)\b(Audio|Sound|Speaker|Microphone|Mute|Volume)\b'                                                                                                = 'Audio'

    # Troubleshooting
    'Regex:(?i)\b(Troubleshoot|Diag|Diagnosis|Fix|Repair|Checkup|Health|SFC|DISM)\b'                                                                            = 'Troubleshooting'

    # Installation
    'Regex:(?i)\b(Install|Setup|Deployment|Sysprep|ImageX|WIM|ISO|Provisioning)\b'                                                                              = 'Installation'

    # Recovery
    'Regex:(?i)\b(Recovery|WinRE|Reset|RestorePoint|SystemRestore|BootRepair)\b'                                                                                = 'Recovery'

    # Utilities
    'Regex:(?i)\b(Util|Utility|Tool|Script|Helper|AdminTool|Sysinternals)\b'                                                                                    = 'Utilities'
  },
  # Optional regex; when set, drop the first derived folder segment if it matches (e.g., to skip a zip root segment)
  [string]$CategoryDropFirstSegmentRegex = ''
)
  
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Cache for performance across a single run (strict-mode safe)
if (-not (Get-Variable -Name HelpCache -Scope Script -EA SilentlyContinue)) { $script:HelpCache = @{} }

# Rose Pine-inspired theme + symbols with variant and Unicode fallbacks
function New-RosePineTheme([string]$variant) {
  switch ($variant) {
    'rose-pine' {
      # Rose Pine: muted with rose/gold accents
      return @{
        Background = [ConsoleColor]::Black
        Foreground = [ConsoleColor]::Gray
        Subtle     = [ConsoleColor]::DarkGray
        Text       = [ConsoleColor]::White
        HeaderBg   = [ConsoleColor]::Black  # Rose accent
        HeaderFg   = [ConsoleColor]::White
        FooterBg   = [ConsoleColor]::Black
        FooterFg   = [ConsoleColor]::DarkCyan
        Border     = [ConsoleColor]::DarkCyan
        CatSelBg   = [ConsoleColor]::Black     # Foam
        CatSelFg   = [ConsoleColor]::Cyan
        CmdSelBg   = [ConsoleColor]::Black      # Rose highlight
        CmdSelFg   = [ConsoleColor]::Cyan
        Accent     = [ConsoleColor]::DarkMagenta
        Highlight  = [ConsoleColor]::Cyan
        Info       = [ConsoleColor]::Cyan
        Warn       = [ConsoleColor]::Yellow       # Gold
        Error      = [ConsoleColor]::DarkRed      # Love
        Preview    = [ConsoleColor]::DarkCyan
        NameColor  = [ConsoleColor]::White
        SectionHdr = [ConsoleColor]::Cyan
      }
    }
    'rose-pine-moon' {
      # Rose Pine Moon: cooler with silver/blue tones
      return @{
        Background = [ConsoleColor]::Black
        Foreground = [ConsoleColor]::Gray
        Subtle     = [ConsoleColor]::DarkGray
        Text       = [ConsoleColor]::White
        HeaderBg   = [ConsoleColor]::DarkBlue     # Moon's cooler tone
        HeaderFg   = [ConsoleColor]::White
        FooterBg   = [ConsoleColor]::DarkBlue
        FooterFg   = [ConsoleColor]::White
        Border     = [ConsoleColor]::DarkBlue
        CatSelBg   = [ConsoleColor]::Blue         # Iris
        CatSelFg   = [ConsoleColor]::White
        CmdSelBg   = [ConsoleColor]::DarkMagenta  # Rose
        CmdSelFg   = [ConsoleColor]::White
        Accent     = [ConsoleColor]::Cyan         # Foam
        Highlight  = [ConsoleColor]::Blue
        Info       = [ConsoleColor]::Cyan
        Warn       = [ConsoleColor]::Yellow
        Error      = [ConsoleColor]::Red
        Preview    = [ConsoleColor]::DarkCyan
        NameColor  = [ConsoleColor]::White
        SectionHdr = [ConsoleColor]::Blue
      }
    }
    'rose-pine-dawn' {
      # Rose Pine Dawn: warm light theme
      return @{
        Background = [ConsoleColor]::White        # Light bg
        Foreground = [ConsoleColor]::DarkGray
        Subtle     = [ConsoleColor]::Gray
        Text       = [ConsoleColor]::Black
        HeaderBg   = [ConsoleColor]::Magenta      # Rose for dawn
        HeaderFg   = [ConsoleColor]::White
        FooterBg   = [ConsoleColor]::Magenta
        FooterFg   = [ConsoleColor]::White
        Border     = [ConsoleColor]::DarkMagenta
        CatSelBg   = [ConsoleColor]::DarkMagenta
        CatSelFg   = [ConsoleColor]::White
        CmdSelBg   = [ConsoleColor]::DarkCyan
        CmdSelFg   = [ConsoleColor]::Black
        Accent     = [ConsoleColor]::DarkCyan
        Highlight  = [ConsoleColor]::Magenta
        Info       = [ConsoleColor]::DarkCyan
        Warn       = [ConsoleColor]::DarkYellow
        Error      = [ConsoleColor]::DarkRed
        Preview    = [ConsoleColor]::DarkCyan
        NameColor  = [ConsoleColor]::Black
        SectionHdr = [ConsoleColor]::DarkMagenta
      }
    }
    'rose-pine-contrast' {
      # High contrast for accessibility
      return @{
        Background = [ConsoleColor]::Black
        Foreground = [ConsoleColor]::White
        Subtle     = [ConsoleColor]::Gray
        Text       = [ConsoleColor]::White
        HeaderBg   = [ConsoleColor]::Magenta
        HeaderFg   = [ConsoleColor]::White
        FooterBg   = [ConsoleColor]::Magenta
        FooterFg   = [ConsoleColor]::White
        Border     = [ConsoleColor]::White
        CatSelBg   = [ConsoleColor]::Cyan
        CatSelFg   = [ConsoleColor]::Black
        CmdSelBg   = [ConsoleColor]::Magenta
        CmdSelFg   = [ConsoleColor]::White
        Accent     = [ConsoleColor]::Cyan
        Highlight  = [ConsoleColor]::Magenta
        Info       = [ConsoleColor]::Cyan
        Warn       = [ConsoleColor]::Yellow
        Error      = [ConsoleColor]::Red
        Preview    = [ConsoleColor]::Cyan
        NameColor  = [ConsoleColor]::White
        SectionHdr = [ConsoleColor]::Magenta
      }
    }
    'dracula' {
      # Dracula: purple bg, cyan/pink/green accents
      return @{
        Background = [ConsoleColor]::Black        # Can't get true purple bg
        Foreground = [ConsoleColor]::White
        Subtle     = [ConsoleColor]::DarkGray
        Text       = [ConsoleColor]::White
        HeaderBg   = [ConsoleColor]::DarkMagenta  # Purple
        HeaderFg   = [ConsoleColor]::White
        FooterBg   = [ConsoleColor]::DarkMagenta
        FooterFg   = [ConsoleColor]::White
        Border     = [ConsoleColor]::DarkMagenta
        CatSelBg   = [ConsoleColor]::Magenta      # Pink
        CatSelFg   = [ConsoleColor]::Black
        CmdSelBg   = [ConsoleColor]::Cyan         # Cyan
        CmdSelFg   = [ConsoleColor]::Black
        Accent     = [ConsoleColor]::Cyan
        Highlight  = [ConsoleColor]::Magenta
        Info       = [ConsoleColor]::Cyan
        Warn       = [ConsoleColor]::Yellow
        Error      = [ConsoleColor]::Red
        Preview    = [ConsoleColor]::Green        # Green comment
        NameColor  = [ConsoleColor]::White
        SectionHdr = [ConsoleColor]::Magenta
      }
    }
    'nord' {
      # Nord: arctic blue-gray palette
      return @{
        Background = [ConsoleColor]::Black
        Foreground = [ConsoleColor]::White
        Subtle     = [ConsoleColor]::DarkGray
        Text       = [ConsoleColor]::White
        HeaderBg   = [ConsoleColor]::DarkBlue     # Polar night
        HeaderFg   = [ConsoleColor]::White
        FooterBg   = [ConsoleColor]::DarkBlue
        FooterFg   = [ConsoleColor]::White
        Border     = [ConsoleColor]::DarkBlue
        CatSelBg   = [ConsoleColor]::Blue         # Frost
        CatSelFg   = [ConsoleColor]::White
        CmdSelBg   = [ConsoleColor]::DarkCyan     # Aurora
        CmdSelFg   = [ConsoleColor]::White
        Accent     = [ConsoleColor]::DarkCyan
        Highlight  = [ConsoleColor]::Blue
        Info       = [ConsoleColor]::DarkCyan
        Warn       = [ConsoleColor]::DarkYellow   # Aurora warm
        Error      = [ConsoleColor]::DarkRed      # Aurora red
        Preview    = [ConsoleColor]::DarkGreen    # Aurora green
        NameColor  = [ConsoleColor]::White
        SectionHdr = [ConsoleColor]::Blue
      }
    }
    'one-dark' {
      # Atom One Dark: purple/blue/cyan with warm accents
      return @{
        Background = [ConsoleColor]::Black
        Foreground = [ConsoleColor]::Gray
        Subtle     = [ConsoleColor]::DarkGray
        Text       = [ConsoleColor]::White
        HeaderBg   = [ConsoleColor]::DarkMagenta  # Purple tones
        HeaderFg   = [ConsoleColor]::White
        FooterBg   = [ConsoleColor]::DarkMagenta
        FooterFg   = [ConsoleColor]::White
        Border     = [ConsoleColor]::DarkGray
        CatSelBg   = [ConsoleColor]::Blue
        CatSelFg   = [ConsoleColor]::White
        CmdSelBg   = [ConsoleColor]::DarkCyan
        CmdSelFg   = [ConsoleColor]::White
        Accent     = [ConsoleColor]::Cyan
        Highlight  = [ConsoleColor]::Blue
        Info       = [ConsoleColor]::Cyan
        Warn       = [ConsoleColor]::Yellow       # Orange-ish
        Error      = [ConsoleColor]::Red
        Preview    = [ConsoleColor]::Green        # Green syntax
        NameColor  = [ConsoleColor]::White
        SectionHdr = [ConsoleColor]::Blue
      }
    }
    'tokyo-night' {
      # Tokyo Night: deep blues with neon cyan/purple
      return @{
        Background = [ConsoleColor]::Black
        Foreground = [ConsoleColor]::Gray
        Subtle     = [ConsoleColor]::DarkGray
        Text       = [ConsoleColor]::White
        HeaderBg   = [ConsoleColor]::DarkBlue     # Deep blue
        HeaderFg   = [ConsoleColor]::Cyan         # Neon accent
        FooterBg   = [ConsoleColor]::DarkBlue
        FooterFg   = [ConsoleColor]::Cyan
        Border     = [ConsoleColor]::DarkBlue
        CatSelBg   = [ConsoleColor]::Blue         # Bright blue
        CatSelFg   = [ConsoleColor]::White
        CmdSelBg   = [ConsoleColor]::Magenta      # Purple accent
        CmdSelFg   = [ConsoleColor]::White
        Accent     = [ConsoleColor]::Cyan         # Neon cyan
        Highlight  = [ConsoleColor]::Magenta      # Purple
        Info       = [ConsoleColor]::Cyan
        Warn       = [ConsoleColor]::Yellow
        Error      = [ConsoleColor]::Red
        Preview    = [ConsoleColor]::Cyan
        NameColor  = [ConsoleColor]::White
        SectionHdr = [ConsoleColor]::Cyan
      }
    }
    'gruvbox' {
      # Gruvbox: warm retro with orange/yellow accents
      return @{
        Background = [ConsoleColor]::Black
        Foreground = [ConsoleColor]::Gray
        Subtle     = [ConsoleColor]::DarkGray
        Text       = [ConsoleColor]::White
        HeaderBg   = [ConsoleColor]::DarkYellow   # Orange-ish
        HeaderFg   = [ConsoleColor]::Black
        FooterBg   = [ConsoleColor]::DarkYellow
        FooterFg   = [ConsoleColor]::Black
        Border     = [ConsoleColor]::DarkYellow
        CatSelBg   = [ConsoleColor]::Yellow       # Bright yellow
        CatSelFg   = [ConsoleColor]::Black
        CmdSelBg   = [ConsoleColor]::DarkGreen    # Aqua/green
        CmdSelFg   = [ConsoleColor]::White
        Accent     = [ConsoleColor]::Green
        Highlight  = [ConsoleColor]::Yellow
        Info       = [ConsoleColor]::DarkCyan
        Warn       = [ConsoleColor]::Yellow
        Error      = [ConsoleColor]::DarkRed
        Preview    = [ConsoleColor]::Green
        NameColor  = [ConsoleColor]::White
        SectionHdr = [ConsoleColor]::Yellow
      }
    }
    'catppuccin-mocha' {
      # Catppuccin Mocha: soft pastels on dark
      return @{
        Background = [ConsoleColor]::Black
        Foreground = [ConsoleColor]::White
        Subtle     = [ConsoleColor]::DarkGray
        Text       = [ConsoleColor]::White
        HeaderBg   = [ConsoleColor]::DarkMagenta  # Mauve
        HeaderFg   = [ConsoleColor]::White
        FooterBg   = [ConsoleColor]::DarkMagenta
        FooterFg   = [ConsoleColor]::White
        Border     = [ConsoleColor]::DarkGray
        CatSelBg   = [ConsoleColor]::Magenta      # Pink
        CatSelFg   = [ConsoleColor]::Black
        CmdSelBg   = [ConsoleColor]::Blue         # Blue
        CmdSelFg   = [ConsoleColor]::White
        Accent     = [ConsoleColor]::Cyan         # Sky
        Highlight  = [ConsoleColor]::Magenta
        Info       = [ConsoleColor]::Cyan
        Warn       = [ConsoleColor]::Yellow       # Peach
        Error      = [ConsoleColor]::Red
        Preview    = [ConsoleColor]::Green        # Green
        NameColor  = [ConsoleColor]::White
        SectionHdr = [ConsoleColor]::Blue
      }
    }
    'solarized-dark' {
      # Solarized Dark: blue-green-yellow balance
      return @{
        Background = [ConsoleColor]::Black
        Foreground = [ConsoleColor]::Gray
        Subtle     = [ConsoleColor]::DarkGray
        Text       = [ConsoleColor]::Gray         # Base0
        HeaderBg   = [ConsoleColor]::DarkCyan     # Cyan
        HeaderFg   = [ConsoleColor]::White
        FooterBg   = [ConsoleColor]::DarkCyan
        FooterFg   = [ConsoleColor]::White
        Border     = [ConsoleColor]::DarkGray
        CatSelBg   = [ConsoleColor]::DarkBlue     # Blue
        CatSelFg   = [ConsoleColor]::White
        CmdSelBg   = [ConsoleColor]::DarkGreen    # Green
        CmdSelFg   = [ConsoleColor]::White
        Accent     = [ConsoleColor]::DarkCyan
        Highlight  = [ConsoleColor]::Blue
        Info       = [ConsoleColor]::DarkCyan
        Warn       = [ConsoleColor]::DarkYellow   # Orange
        Error      = [ConsoleColor]::DarkRed
        Preview    = [ConsoleColor]::Green
        NameColor  = [ConsoleColor]::Gray
        SectionHdr = [ConsoleColor]::Blue
      }
    }
    default {
      # Default Windows Terminal style
      return @{
        Background = [ConsoleColor]::Black
        Foreground = [ConsoleColor]::Gray
        Subtle     = [ConsoleColor]::DarkGray
        Text       = [ConsoleColor]::White
        HeaderBg   = [ConsoleColor]::DarkBlue
        HeaderFg   = [ConsoleColor]::White
        FooterBg   = [ConsoleColor]::DarkBlue
        FooterFg   = [ConsoleColor]::White
        Border     = [ConsoleColor]::DarkGray
        CatSelBg   = [ConsoleColor]::Blue
        CatSelFg   = [ConsoleColor]::White
        CmdSelBg   = [ConsoleColor]::DarkCyan
        CmdSelFg   = [ConsoleColor]::White
        Accent     = [ConsoleColor]::Cyan
        Highlight  = [ConsoleColor]::Blue
        Info       = [ConsoleColor]::Cyan
        Warn       = [ConsoleColor]::Yellow
        Error      = [ConsoleColor]::Red
        Preview    = [ConsoleColor]::Cyan
        NameColor  = [ConsoleColor]::White
        SectionHdr = [ConsoleColor]::Blue
      }
    }
  }
}

# Ensure high-contrast foreground/background pairs
function Ensure-ThemeContrast($t) {
  # # For dark backgrounds, prefer white text; for light-ish backgrounds, prefer black text
  # $darkBgs = @([ConsoleColor]::Black, [ConsoleColor]::DarkBlue, [ConsoleColor]::DarkGreen, [ConsoleColor]::DarkCyan, [ConsoleColor]::DarkRed, [ConsoleColor]::DarkMagenta, [ConsoleColor]::DarkYellow, [ConsoleColor]::DarkGray)
  # if ($darkBgs -contains $t.Background) {
  #   if ($t.Foreground -eq [ConsoleColor]::Black) { $t.Foreground = [ConsoleColor]::Gray }
  #   if ($t.Text -eq [ConsoleColor]::Black) { $t.Text = [ConsoleColor]::White }
  #   if ($t.SectionHdr -eq [ConsoleColor]::DarkYellow) { $t.SectionHdr = [ConsoleColor]::Yellow }
  # }
  # else {
  #   # Light backgrounds
  #   if ($t.Foreground -eq [ConsoleColor]::White) { $t.Foreground = [ConsoleColor]::Black }
  #   if ($t.Text -eq [ConsoleColor]::White) { $t.Text = [ConsoleColor]::Black }
  # }
  # # Selection pairs: ensure strong contrast
  # if ($t.CatSelBg -eq $t.CatSelFg) { $t.CatSelFg = [ConsoleColor]::Black }
  # if ($t.CmdSelBg -eq $t.CmdSelFg) { $t.CmdSelFg = [ConsoleColor]::Black }
  return $t
}

function New-ThemeSymbols([bool]$useUnicode) {
  if ($useUnicode) {
    return @{
      Box     = $(if ($ThemeStyle -eq 'minimal') { @{ TL = ' '; TR = ' '; BL = ' '; BR = ' '; H = ' '; V = ' ' } } else { @{ TL = "$([char]0x256D)"; TR = "$([char]0x256E)"; BL = "$([char]0x2570)"; BR = "$([char]0x256F)"; H = "$([char]0x2500)"; V = "$([char]0x2502)" } })
      Up      = "$([char]0x25B2)"
      Down    = "$([char]0x25BC)"
      Bullet  = $(if ($ThemeStyle -eq 'minimal') { '  ' + [char]0x2219 + ' ' } else { ' ' + [char]0x2022 + ' ' })
      Sparkle = "$([char]0x2726)"
    }
  }
  else {
    return @{
      Box     = $(if ($ThemeStyle -eq 'minimal') { @{ TL = ' '; TR = ' '; BL = ' '; BR = ' '; H = ' '; V = ' ' } } else { @{ TL = '+'; TR = '+'; BL = '+'; BR = '+'; H = '-'; V = '|' } })
      Up      = "$([char]0x005E)"
      Down    = "$([char]0x0076)"
      Bullet  = $(if ($ThemeStyle -eq 'minimal') { '   ' } else { ' - ' })
      Sparkle = '*'
    }
  }
}

$script:Theme = Ensure-ThemeContrast (New-RosePineTheme -variant $ThemeVariant)
$script:Symbols = New-ThemeSymbols -useUnicode:(!$NoUnicode)

function Write-Info($m) { Write-Host $m -ForegroundColor $script:Theme.Info }
function Write-Warn($m) { Write-Host $m -ForegroundColor $script:Theme.Warn }


function Safe-GetHelp([string]$name) {
  if ($script:HelpCache.ContainsKey($name)) { return $script:HelpCache[$name] }
  $prevProg = $ProgressPreference
  try {
    $ProgressPreference = 'SilentlyContinue'
    $h = (Get-Help -Name $name -EA Stop)
    $script:HelpCache[$name] = $h
    return $h
  }
  catch {
    $script:HelpCache[$name] = $null
    return $null
  }
  finally {
    $ProgressPreference = $prevProg
  }
}

 

function Parse-GitHubRepoInfo([string]$url, [string]$branchDefault = 'main') {
  if (-not $url) { return $null }
  $owner = $null; $repo = $null; $branch = $null
  if ($url -match 'githubusercontent\.com/([^/]+)/([^/]+)/([^/]+)/') {
    $owner = $Matches[1]; $repo = $Matches[2]; $branch = $Matches[3]
  }
  elseif ($url -match 'github\.com/([^/]+)/([^/]+)(?:/(?:tree|blob)/([^/]+)/)?') {
    $owner = $Matches[1]; $repo = $Matches[2]; if ($Matches[3]) { $branch = $Matches[3] }
  }
  if (-not $owner -or -not $repo) { return $null }
  if (-not $branch) { $branch = $branchDefault }
  [pscustomobject]@{ Owner = $owner; Repo = $repo; Branch = $branch }
}

function Get-InstallBase([string]$preferred) {
  if ($preferred) { return $preferred }
  if ($env:USERNAME -eq 'SYSTEM') { return (Join-Path $env:SystemDrive 'Tools') }
  return (Join-Path $env:LOCALAPPDATA 'VTS')
}

function Resolve-FunctionsPath([string]$inputPath) {
  # Choose a sane default when running from memory (IEX) where $PSScriptRoot is empty
  $candidates = @()
  if ($inputPath -and -not [string]::IsNullOrWhiteSpace($inputPath)) { $candidates += , $inputPath }
  if ($PSScriptRoot -and -not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $candidates += , $PSScriptRoot }
  try { if ($PWD -and $PWD.Path) { $candidates += , $PWD.Path } } catch { }

  foreach ($p in $candidates) {
    try {
      if ($p -and -not [string]::IsNullOrWhiteSpace($p)) {
        $full = [IO.Path]::GetFullPath($p)
        if (Test-Path -LiteralPath $full) {
          # Prefer a 'functions' subfolder if present
          $funcDir = Join-Path $full 'functions'
          if (Test-Path -LiteralPath $funcDir) { return $funcDir }
          return $full
        }
      }
    }
    catch { }
  }
  # Fallback to current directory string if all else fails
  try { return $PWD.Path } catch { return '' }
}

function Install-And-RunFromRepo([string]$repoUrl, [string]$branch, [string]$startRelPath, [string]$installBase, [switch]$reinstall) {
  $info = Parse-GitHubRepoInfo -url $repoUrl -branchDefault $branch
  if (-not $info) { throw "Unable to parse GitHub repo from URL: $repoUrl" }
  $base = Get-InstallBase -preferred $installBase
  if (-not (Test-Path $base)) { New-Item -ItemType Directory -Path $base -Force | Out-Null }
  $zipUrl = "https://github.com/$($info.Owner)/$($info.Repo)/archive/refs/heads/$($info.Branch).zip"
  $zipPath = Join-Path $env:TEMP "$($info.Repo)-$($info.Branch).zip"
  $effectiveBranch = $info.Branch
  $repoDir = Join-Path $base "$($info.Repo)-$effectiveBranch"
  if ($reinstall -and (Test-Path $repoDir)) {
    Write-Warn "Removing existing folder: $repoDir"
    Remove-Item -Recurse -Force -Path $repoDir -EA SilentlyContinue
  }
  $downloaded = $false
  $prevTLS = $null
  try { $prevTLS = [Net.ServicePointManager]::SecurityProtocol } catch { }
  try {
    try { [Net.ServicePointManager]::SecurityProtocol = $prevTLS -bor [Net.SecurityProtocolType]::Tls12 } catch { }
    foreach ($b in @($info.Branch, $(if ($info.Branch -ieq 'main') { 'master' } elseif ($info.Branch -ieq 'master') { 'main' } else { $null }))) {
      if (-not $b) { continue }
      $tryZipUrl = "https://github.com/$($info.Owner)/$($info.Repo)/archive/refs/heads/$b.zip"
      try {
        Write-Info "Downloading $tryZipUrl ..."
        Invoke-WebRequest -Uri $tryZipUrl -UseBasicParsing -OutFile $zipPath
        $downloaded = $true
        $effectiveBranch = $b
        break
      }
      catch {
        Write-Warn ("Download failed for branch '$b': " + $_.Exception.Message)
        continue
      }
    }
  }
  finally {
    try { if ($null -ne $prevTLS) { [Net.ServicePointManager]::SecurityProtocol = $prevTLS } } catch { }
  }
  if (-not $downloaded) { throw "Failed to download repository ZIP from $zipUrl" }
  try {
    if (Test-Path $repoDir) {
      Write-Warn "Removing existing folder: $repoDir"
      Remove-Item -Recurse -Force -Path $repoDir -EA SilentlyContinue
    }
    Write-Info "Extracting to $base ..."
    try { Expand-Archive -Path $zipPath -DestinationPath $base -Force }
    catch {
      try { Import-Module Microsoft.PowerShell.Archive -EA Stop; Expand-Archive -Path $zipPath -DestinationPath $base -Force }
      catch { throw }
    }
  }
  finally {
    if (Test-Path $zipPath) { Remove-Item -Force $zipPath -EA SilentlyContinue }
  }
  $repoDir = Join-Path $base "$($info.Repo)-$effectiveBranch"
  $startScript = Join-Path $repoDir $startRelPath
  if ((Test-Path -LiteralPath $startScript) -and (-not $BootstrapScanOnly)) {
    Write-Info "Launching $startRelPath from $repoDir ..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $startScript
    return
  }
  else {
    Write-Warn "Scanning repo for commands (no auto-execution). Path: $repoDir"
    try {
      # Auto-drop the zip's top-level folder (e.g., repo-branch)
      $prevDrop = $script:CategoryDropFirstSegmentRegex
      $script:CategoryDropFirstSegmentRegex = "^(?i)$([Regex]::Escape("$($info.Repo)-$effectiveBranch"))$"
      $model = Build-Model -functionsPath $repoDir -CategoryRoot $repoDir
      Run-TUI -model $model
      return
    }
    catch {
      throw "Start script not found and embedded TUI failed: $($_.Exception.Message)"
    }
    finally {
      $script:CategoryDropFirstSegmentRegex = $prevDrop
    }
  }
}

function Ensure-Console {
  if (-not ($Host.Name -match 'ConsoleHost|Visual Studio Code')) {
    Write-Host 'This script is designed for an interactive console host.' -ForegroundColor $script:Theme.Warn
  }
  try { [void][Console]::CursorVisible } catch {
    throw "This host doesn't expose [Console]. Try running in Windows Terminal/PowerShell console."
  }
}

function Safe-Substring([string]$s, [int]$len) {
  if ([string]::IsNullOrEmpty($s)) { return '' }
  if ($s.Length -le $len) { return $s }
  return ($s.Substring(0, [Math]::Max(0, $len - 3)) + '...')
}

function TruncatePad([string]$s, [int]$width) {
  $s = ($s -replace '\s+', ' ').Trim()
  $t = Safe-Substring $s $width
  return $t + (' ' * [Math]::Max(0, $width - ($t).Length))
}

function Read-Key { [Console]::ReadKey($true) }

function Draw-Line([int]$y, [int]$x, [int]$width, [ConsoleColor]$fg, [string]$ch = ' ') {
  try {
    if ($width -le 0) { return }
    [Console]::SetCursorPosition([Math]::Max(0, $x), [Math]::Max(0, $y))
    $s = $ch * $width
    Write-Host $s -NoNewline -ForegroundColor $fg
  }
  catch { }
}

function Draw-VertLine([int]$top, [int]$x, [int]$height, [ConsoleColor]$fg, [string]$ch = '|') {
  try {
    if ($height -le 0) { return }
    for ($yy = 0; $yy -lt $height; $yy++) {
      [Console]::SetCursorPosition([Math]::Max(0, $x), [Math]::Max(0, $top + $yy))
      Write-Host $ch -NoNewline -ForegroundColor $fg
    }
  }
  catch { }
}

function Write-At([int]$y, [int]$x, [string]$text, [ConsoleColor]$fg, [Nullable[ConsoleColor]]$bg = $null, [switch]$NoNewline) {
  try {
    if ($y -ge 0 -and $x -ge 0) { [Console]::SetCursorPosition($x, $y) }
    if ($bg -ne $null) { [Console]::BackgroundColor = $bg }
    Write-Host $text -ForegroundColor $fg -NoNewline:$NoNewline
    if ($bg -ne $null) { [Console]::ResetColor() }
  }
  catch { }
}

function Fill-Rect([int]$top, [int]$left, [int]$height, [int]$width, [ConsoleColor]$bg) {
  try {
    if ($height -le 0 -or $width -le 0) { return }
    for ($yy = 0; $yy -lt $height; $yy++) {
      [Console]::SetCursorPosition([Math]::Max(0, $left), [Math]::Max(0, $top + $yy))
      [Console]::BackgroundColor = $bg
      [Console]::Write((' ' * [Math]::Max(0, $width)))
      [Console]::ResetColor()
    }
  }
  catch { }
}

function Box([int]$top, [int]$left, [int]$height, [int]$width, [string]$title, [ConsoleColor]$fg = $(if ($script:Theme) { $script:Theme.Border } else { [ConsoleColor]::Gray })) {
  if ($ThemeStyle -eq 'minimal') {
    # No borders; render a subtle title line only
    if ($title) {
      $t = TruncatePad $title ($width - 2)
  Write-At $top ($left + 1) $t $fg $script:Theme.Background
    }
    return
  }
  $right = $left + $width - 1
  $bottom = $top + $height - 1
  # Top
  Write-At $top $left $script:Symbols.Box.TL $fg
  Draw-Line $top ($left + 1) ($width - 2) $fg $script:Symbols.Box.H
  Write-At $top $right $script:Symbols.Box.TR $fg
  # Sides
  for ($y = $top + 1; $y -lt $bottom; $y++) {
    Write-At $y $left $script:Symbols.Box.V $fg
    Write-At $y $right $script:Symbols.Box.V $fg
  }
  # Bottom
  Write-At $bottom $left $script:Symbols.Box.BL $fg
  Draw-Line $bottom ($left + 1) ($width - 2) $fg $script:Symbols.Box.H
  Write-At $bottom $right $script:Symbols.Box.BR $fg
  if ($title) {
    $titleClean = Safe-Substring $title ($width - 4)
    Write-At $top ($left + 2) " $titleClean " $fg
  }
}

function Wrap([string]$t, [int]$w) {
  if (-not $t) { return @() }
  $t = ($t -replace '\s+', ' ').Trim()
  $out = @()
  while ($t.Length -gt $w) {
    $break = $t.LastIndexOf(' ', [Math]::Min($w, $t.Length - 1))
    if ($break -lt 1) { $break = $w }
    $out += $t.Substring(0, $break)
    $t = $t.Substring($break).Trim()
  }
  if ($t) { $out += $t }
  return $out
}

function Wrap-Block([string]$text, [int]$w) {
  if (-not $text) { return @() }
  $nl = ($text -replace "\r\n", "`n") -replace "\r", "`n"
  $paras = $nl -split "`n`n+"  # split on blank lines
  $out = @()
  foreach ($p in $paras) {
    $p2 = $p.Trim()
    if (-not $p2) { $out += ''; continue }
    $out += (Wrap $p2 $w)
    $out += ''
  }
  if ($out.Count -gt 0 -and $out[-1] -eq '') { $out = $out[0..($out.Count - 2)] }
  return $out
}

function Wrap-Bullet([string]$text, [int]$w, [string]$prefix = ' - ', [string]$cont = '   ') {
  if (-not $text) { return @() }
  $inner = [Math]::Max(8, $w - $prefix.Length)
  $wrapped = @(Wrap $text $inner)  # force array to avoid .Count on scalars
  $out = @()
  $first = $true
  foreach ($seg in $wrapped) {
    $out += ($(if ($first) { $prefix }else { $cont }) + $seg)
    $first = $false
  }
  return $out
}

# Wrap a single line without collapsing whitespace; preserves spaces and breaks hard at width
function Wrap-LineLiteral([string]$text, [int]$w) {
  if ($null -eq $text) { return @('') }
  if ($w -le 0) { return @($text) }
  $t = "$text"
  $out = @()
  $i = 0
  while ($i -lt $t.Length) {
    $len = [Math]::Min($w, $t.Length - $i)
    $out += $t.Substring($i, $len)
    $i += $len
  }
  if ($out.Count -eq 0) { return @('') }
  return $out
}

function Shorten-UrlLabel([string]$s) {
  if (-not $s) { return $s }
  $t = "$s".Trim()
  # Drop anchors and queries
  $t = ($t -replace '#.*$', '')
  $t = ($t -replace '\?.*$', '')
  # GitHub special-case -> owner/repo@branch
  $gh = $null
  try { $gh = Parse-GitHubRepoInfo -url $t -branchDefault $Branch } catch { $gh = $null }
  if ($gh) {
  $lbl = "$($gh.Owner)/$($gh.Repo)"
  if ($gh.Branch) { $lbl += "@$($gh.Branch)" }
  return $lbl
  }
  # Generic: host[/first-segment]
  if ($t -match '^(?i)(https?://|www\.)') {
    try {
      if ($t -notmatch '^(?i)https?://') { $t = 'https://' + $t }
      $u = [Uri]$t
      $hostName = $u.Host
      $path = $u.AbsolutePath.Trim('/')
      if ($path) { return ("{0}/{1}" -f $hostName, ($path.Split('/')[0])) }
      return $hostName
    }
    catch { return $s }
  }
  return $s
}

# Heuristically clean long or noisy category texts (from .LINK LinkText, etc.)
function Sanitize-CategoryText([string]$text) {
  if ([string]::IsNullOrWhiteSpace($text)) { return $text }
  $t = ("$text" -replace '\s+', ' ').Trim()
  # Prefer left side of common separators
  foreach ($sep in @(' - ', ' — ', ':', ' | ')) {
    $idx = $t.IndexOf($sep)
    if ($idx -gt 0) { $t = $t.Substring(0, $idx).Trim(); break }
  }
  # If still too long, clamp by words/length
  $words = @($t.Split(' '))
  if ($words.Count -gt 5 -or $t.Length -gt 40) {
    $t = ($words[0..([Math]::Min(3, $words.Count - 1))] -join ' ')
  }
  # Drop trailing filler words
  $t = ($t -replace '\b(for|with|using|via|and|to|on)$', '').Trim()
  return $t
}

# Convert a compact owner/repo[@branch] or host/segment into a friendlier label
function Format-FriendlyCategoryLabel([string]$label, [string]$style) {
  if ([string]::IsNullOrWhiteSpace($label)) { return $label }
  $textInfo = [Globalization.CultureInfo]::InvariantCulture.TextInfo
  $fmtToken = {
    param($tok)
    # Only title-case tokens that are plain lower/number/kebab/underscore/dot to avoid mangling CamelCase
    if ($tok -match '^[a-z0-9._-]+$') {
      $clean = ($tok -replace '[-_\.]+', ' ')
      return $textInfo.ToTitleCase($clean.ToLowerInvariant()).Trim()
    }
    return $tok
  }
  $s = "$label".Trim()
  # owner/repo[@branch]
  if ($s -match '^([^/@]+)/([^/@]+)(?:@([^@]+))?$') {
    $owner = $Matches[1]; $repo = $Matches[2]; $branch = $Matches[3]
    switch ($style) {
      'Repo' { return & $fmtToken $repo }
      'OwnerRepo' { return ((& $fmtToken $owner) + '/' + (& $fmtToken $repo)) }
      default {
        # Full
        $core = ((& $fmtToken $owner) + '/' + (& $fmtToken $repo))
        if ($branch) { return ($core + '@' + (& $fmtToken $branch)) }
        return $core
      }
    }
  }
  # host/segment
  if ($s -match '^[^/]+/([^/]+)$') {
    $seg = $Matches[1]
    if ($style -eq 'Repo') { return & $fmtToken $seg }
    # For OwnerRepo/Full, keep original host/seg but prettify segment only
    $hostPart = $s.Substring(0, $s.Length - $seg.Length - 1)
    return "$hostPart/" + (& $fmtToken $seg)
  }
  # Fallback: de-kebab and title-case lightly
  return (& $fmtToken $s)
}

function Normalize-Category([object]$c) {
  # Normalize a string or list of strings using the override map; otherwise return cleaned label
  if ($c -is [System.Collections.IEnumerable] -and -not ($c -is [string])) {
    foreach ($e in $c) {
      $n = Normalize-Category $e
      if ($n -and $n -ne 'Uncategorized') { return $n }
    }
    return 'Uncategorized'
  }
  $s = ''
  try { $s = if ($null -eq $c) { '' } else { [string]$c } } catch { $s = '' }
  $s = ($s -replace '\s+', ' ').Trim()
  if ([string]::IsNullOrWhiteSpace($s)) { return 'Uncategorized' }
  if ($s -match '^(?i)uncategorized$') { return 'Uncategorized' }
  try {
    if ($CategoryOverrideMap -and $CategoryOverrideMap.Count -gt 0) {
      foreach ($k in $CategoryOverrideMap.Keys) {
        $key = "$k"; $val = "$($CategoryOverrideMap[$k])"
        if ($key -like 'Regex:*') {
          $pat = $key.Substring(6)
          if ($s -match $pat) { return $val }
        }
        elseif ($s -ieq $key) { return $val }
      }
      foreach ($v in $CategoryOverrideMap.Values) { if ("$v" -ieq $s) { return "$v" } }
    }
  }
  catch { }
  # No override matched; return cleaned friendly label
  return (Format-FriendlyCategoryLabel -label (Sanitize-CategoryText $s) -style $CategoryLabelStyle)
}

function Prompt-String([string]$label, [string]$default = '') {
  $prompt = if ($default) { "$label [$default]" } else { $label }
  Read-Host $prompt
}

function Try-Cast([string]$value, [string]$typeName, [ref]$casted) {
  if ($null -eq $value) { $casted.Value = $null; return $true }
  if ($value -eq '' ) { $casted.Value = $null; return $true }
  $tn = "$typeName"
  try {
    # Arrays: split by comma and cast elements to inner type
    if ($tn -match '^((System\.)?\w+)\[\]$') {
      $elemType = $Matches[1]
      $parts = $value -split '\s*,\s*'
      $arr = @()
      foreach ($part in $parts) {
        $tmp = $null
        if (-not (Try-Cast $part $elemType ([ref]$tmp))) { $tmp = $part }
        $arr += , $tmp
      }
      $casted.Value = $arr
      return $true
    }
    if ($tn -match '^(System\.)?Int16$') { $casted.Value = [int16]$value; return $true }
    elseif ($tn -match '^(System\.)?Int32$') { $casted.Value = [int]$value; return $true }
    elseif ($tn -match '^(System\.)?Int64$') { $casted.Value = [long]$value; return $true }
    elseif ($tn -match '^(System\.)?UInt16$') { $casted.Value = [uint16]$value; return $true }
    elseif ($tn -match '^(System\.)?UInt32$') { $casted.Value = [uint32]$value; return $true }
    elseif ($tn -match '^(System\.)?UInt64$') { $casted.Value = [uint64]$value; return $true }
    elseif ($tn -match '^(System\.)?Double$') { $casted.Value = [double]$value; return $true }
    elseif ($tn -match '^(System\.)?Single$') { $casted.Value = [single]$value; return $true }
    elseif ($tn -match '^(System\.)?Decimal$') { $casted.Value = [decimal]$value; return $true }
    elseif ($tn -match '^(System\.)?Boolean$') { $casted.Value = [bool]$value; return $true }
    elseif ($tn -match '^(System\.)?Date(Time)?$') { $casted.Value = [datetime]$value; return $true }
    elseif ($tn -match '^(System\.)?String$') { $casted.Value = [string]$value; return $true }
    elseif ($tn -match '^(System\.)?Management\.Automation\.PSCredential$') {
      $cred = Get-Credential -Message "Enter PSCredential for parameter"
      $casted.Value = $cred
      return $true
    }
    else { $casted.Value = $value; return $true }
  }
  catch {
    return $false
  }
}

function Load-ModuleFromFiles([System.IO.FileInfo[]]$files) {
  if (-not $files -or $files.Count -eq 0) { return @() }
  $modules = @()
  foreach ($f in $files) {
    try {
      $tokens = $null; $errors = $null
      $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errors)
      if ($errors -and $errors.Count -gt 0) { continue }
      $funcAsts = $ast.FindAll({ param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
      if (-not $funcAsts -or $funcAsts.Count -eq 0) { continue }
      $funcTexts = @()
      $funcNames = @()
      foreach ($fa in $funcAsts) {
        $name = $fa.Name
        if (-not $name) { continue }
        $funcNames += , $name
        # Use the original function text to preserve attributes and comments
        $funcTexts += , ($fa.Extent.Text)
      }
      if ($funcTexts.Count -eq 0) { continue }
      $code = ($funcTexts -join "`n`n") + "`n`n" + ("Export-ModuleMember -Function @('{0}')" -f (($funcNames -join "','")))
      $modName = ("VTS.Dynamic.{0}" -f ([IO.Path]::GetFileNameWithoutExtension($f.Name)))
      $module = New-Module -Name $modName -ScriptBlock ([ScriptBlock]::Create($code))
      try { Import-Module $module -Force -DisableNameChecking } catch { continue }
      $modules += , $module
    }
    catch { }
  }
  return $modules
}

function Get-RelativePath([string]$base, [string]$full) {
  try {
    $b = [IO.Path]::GetFullPath($base)
    $f = [IO.Path]::GetFullPath($full)
    # Try .NET relative path if available
    try {
      $relTry = [IO.Path]::GetRelativePath($b, $f)
      if ($relTry -and -not ($relTry.StartsWith('..'))) { return ($relTry -replace '\\', '/') }
    }
    catch { }
    # Ensure trailing separator on base for safe prefix comparison
    $sep = [IO.Path]::DirectorySeparatorChar
    if (-not $b.EndsWith("$sep")) { $b = $b + "$sep" }
    if ($f.ToLowerInvariant().StartsWith($b.ToLowerInvariant())) {
      $rel = $f.Substring($b.Length).TrimStart('\\', '/')
      return ($rel -replace '\\', '/')
    }
    # Fallback: URI-based relative path (more robust across cases)
    try {
      $bUri = [Uri]($b.TrimEnd($sep) + "$sep")
      $fUri = [Uri]$f
      $relUri = $bUri.MakeRelativeUri($fUri)
      $rel2 = [Uri]::UnescapeDataString($relUri.ToString())
      if ($rel2) { return ($rel2 -replace '\\', '/') }
    }
    catch { }
  }
  catch { }
  return $full
}

function Resolve-CategoryRoot([string]$path) {
  if (-not (Get-Variable -Name CategoryRootCache -Scope Script -EA SilentlyContinue)) { $script:CategoryRootCache = @{} }
  $keyInput = $path
  try { if (Test-Path -LiteralPath $path -PathType Leaf) { $keyInput = Split-Path -Path $path -Parent } } catch { }
  try { $key = [IO.Path]::GetFullPath($keyInput) } catch { $key = $keyInput }
  try { if ($script:CategoryRootCache.ContainsKey($key)) { return $script:CategoryRootCache[$key] } } catch { }
  try {
    $dir = $path
    if (Test-Path -LiteralPath $path -PathType Leaf) { $dir = Split-Path -Path $path -Parent }
    $cur = [IO.Path]::GetFullPath($dir)
  }
  catch { return $path }
  $maxUp = 12
  for ($i = 0; $i -lt $maxUp -and $cur -and (Test-Path -LiteralPath $cur); $i++) {
    $git = Join-Path $cur '.git'
    $gh = Join-Path $cur '.github'
    $readme = Get-ChildItem -LiteralPath $cur -File -Filter README* -EA SilentlyContinue | Select-Object -First 1
    $license = Get-ChildItem -LiteralPath $cur -File -Filter LICENSE* -EA SilentlyContinue | Select-Object -First 1
    if ((Test-Path -LiteralPath $git -PathType Any -EA SilentlyContinue) -or
      (Test-Path -LiteralPath $gh -PathType Any -EA SilentlyContinue) -or
      ($null -ne $readme) -or ($null -ne $license)) {
      return $cur
    }
    try { $parent = Split-Path -Path $cur -Parent } catch { $parent = $null }
    if (-not $parent -or $parent -eq $cur) { break }
    $cur = $parent
  }
  # If we didn't find a marker, prefer the initial directory if it contains scripts; otherwise nearest with scripts
  try {
    $self = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    $ps1Here = Get-ChildItem -LiteralPath $dir -Filter *.ps1 -File -EA SilentlyContinue | Where-Object { $_.FullName -ne $self } | Select-Object -First 1
    if ($ps1Here) {
      try { $script:CategoryRootCache[$key] = $dir } catch { }
      return $dir
    }
  }
  catch { }
  try { $script:CategoryRootCache[$key] = $dir } catch { }
  return $dir
}

function Derive-Category([string]$base, [System.IO.FileInfo]$file) {
  if (-not (Get-Variable -Name DeriveCategoryCache -Scope Script -EA SilentlyContinue)) { $script:DeriveCategoryCache = @{} }
  $cacheKey = "{0}|{1}" -f (try { [IO.Path]::GetFullPath($base) } catch { $base }), (try { [IO.Path]::GetFullPath($file.DirectoryName) } catch { $file.DirectoryName })
  if ($script:DeriveCategoryCache.ContainsKey($cacheKey)) { return $script:DeriveCategoryCache[$cacheKey] }
  $rel = Get-RelativePath $base $file.DirectoryName
  if ([string]::IsNullOrWhiteSpace($rel)) { return 'Root' }
  $parts = @($rel -split '[\\/]+' | Where-Object { $_ -and $_.Trim() -ne '' })
  if (-not $parts.Count) { return 'Root' }
  # Remove leading ignored folders
  $idx = 0
  while ($idx -lt $parts.Count) {
    $p = "$($parts[$idx])".Trim()
    # Skip drive letters like 'C:' or UNC roots
    if ($p -match '^[A-Za-z]:$' -or $p -eq '' -or $p -eq '\\') { $idx++; continue }
    if ($CategoryIgnoreFolders -and ($CategoryIgnoreFolders | Where-Object { $_ -and ($_ -ieq $p) } | Select-Object -First 1)) {
      $idx++
      continue
    }
    break
  }
  if ($idx -ge $parts.Count) { return 'Root' }
  # Remove file-like segments (extensions) and common noise
  $cleaned = @()
  for ($j = $idx; $j -lt $parts.Count; $j++) {
    $seg = "$($parts[$j])"
    if ($seg -match '^README(\.md|\.txt)?$') { continue }
    if ($seg -match '^(bin|obj|out|build|dist|release|debug)$') { continue }
    if ($seg -match '^v?\d+\.[\d\._-]+$') { continue }
    if ($seg -match '^.+\.(ps1|psm1|dll|exe|json|yaml|yml|md)$') { continue }
    $cleaned += , $seg
  }
  if (-not $cleaned.Count) { return 'Root' }
  # Optionally drop first segment if it matches a provided regex
  if ($CategoryDropFirstSegmentRegex -and $cleaned.Count -gt 0) {
    try { if ("$($cleaned[0])" -match $CategoryDropFirstSegmentRegex) { $cleaned = $cleaned[1..($cleaned.Count - 1)] } } catch { }
  }
  $usable = $cleaned[0..([Math]::Min($cleaned.Count - 1, $CategoryDepth - 1))]
  # Apply rename map (case-insensitive)
  $nice = foreach ($seg in $usable) {
    $key = $seg
    $mapped = $null
    if ($CategoryRenameMap -and ($CategoryRenameMap.Keys | Where-Object { $_ -and ($_ -ieq $key) } | Select-Object -First 1)) {
      foreach ($k in $CategoryRenameMap.Keys) { if ($k -ieq $key) { $mapped = $CategoryRenameMap[$k]; break } }
    }
    $val = if ($mapped) { "$mapped" } else { "$seg" }
    # Prettify each segment
    (Format-FriendlyCategoryLabel -label $val -style 'Repo')
  }
  $label = ($nice -join ' / ')
  if (-not $label) { $label = 'Root' }
  $final = (Normalize-Category $label)
  try { $script:DeriveCategoryCache[$cacheKey] = $final } catch { }
  return $final
}

function Detect-FunctionNameInFile([System.IO.FileInfo]$file) {
  try {
    $lines = Get-Content -Path $file.FullName -TotalCount 200 -EA SilentlyContinue
    if (-not $lines) { return $null }
    foreach ($ln in $lines) {
      if ($ln -match '^\s*function\s+([A-Za-z_][\w-]*)\s*') { return $Matches[1] }
    }
  }
  catch { }
  return $null
}

function Get-CategoryForItem([string]$cmdName, [string]$filePath, [string]$root, [string]$parsedCat, [switch]$DisableHelp) {
  # Resolve help-derived and folder-derived candidates
  $helpCat = $null
  if (-not $DisableHelp -and $parsedCat) {
    $s = ("$parsedCat").Trim()
    if ($s) {
      if ($s -match '^(?i)(https?://|www\.|mailto:|file:)') { $helpCat = (Format-FriendlyCategoryLabel -label (Shorten-UrlLabel $s) -style $CategoryLabelStyle) }
      else { $helpCat = (Format-FriendlyCategoryLabel -label (Sanitize-CategoryText $s) -style $CategoryLabelStyle) }
    }
  }
  $folderCat = $null
  try {
    $catRoot = if ($root) { $root } else { Resolve-CategoryRoot -path $filePath }
    $fileObj = Get-Item -LiteralPath $filePath -EA SilentlyContinue
    if ($fileObj) { $folderCat = Derive-Category -base $catRoot -file $fileObj }
  }
  catch { }

  # Apply precedence: PreferFolderCategory means folder first, else help first
  if ($PreferFolderCategory) {
    $nFolder = if ($folderCat) { Normalize-Category $folderCat } else { 'Uncategorized' }
    if ($nFolder -and $nFolder -ne 'Uncategorized') { return $nFolder }
    $nHelp = if ($helpCat) { Normalize-Category $helpCat } else { 'Uncategorized' }
    if ($nHelp -and $nHelp -ne 'Uncategorized') { return $nHelp }
  }
  else {
    $nHelp = if ($helpCat) { Normalize-Category $helpCat } else { 'Uncategorized' }
    if ($nHelp -and $nHelp -ne 'Uncategorized') { return $nHelp }
    $nFolder = if ($folderCat) { Normalize-Category $folderCat } else { 'Uncategorized' }
    if ($nFolder -and $nFolder -ne 'Uncategorized') { return $nFolder }
  }
  # Name-based override
  $candidates = @()
  if ($cmdName) { $candidates += , $cmdName }
  try { $base = [IO.Path]::GetFileNameWithoutExtension($filePath); if ($base) { $candidates += , $base } } catch { }
  foreach ($cand in $candidates) {
    $n = Normalize-Category $cand
    if ($n -and $n -ne 'Uncategorized') { return $n }
  }
  # Last resort: root folder category
  try {
    $catRoot = if ($root) { $root } else { Resolve-CategoryRoot -path $filePath }
    $fileObj = Get-Item -LiteralPath $filePath -EA SilentlyContinue
    if ($fileObj) { return (Derive-Category -base $catRoot -file $fileObj) }
  }
  catch { }
  return 'Uncategorized'
}


function Parse-HelpFromFile([string]$filePath) {
  $syn = ''
  $desc = ''
  $examples = @()
  $category = 'Uncategorized'
  $lines = Get-Content -Path $filePath -EA SilentlyContinue -TotalCount 400
  if (-not $lines) { return [pscustomobject]@{ Synopsis = ''; Description = ''; Examples = @(); Category = 'Uncategorized' } }
  $inHelp = $false
  $curTag = ''
  $buf = @()

  foreach ($raw in $lines) {
    $line = "$raw"
    if (-not $inHelp) {
      if ($line -match '^\s*<#') { $inHelp = $true }
      continue
    }
    if ($line -match '#>') {
      # flush
      $text = ($buf -join " `n").Trim()
      if ($curTag -eq 'SYNOPSIS' -and $text) { $syn = $text }
      elseif ($curTag -eq 'DESCRIPTION' -and $text) { $desc = $text }
      elseif ($curTag -eq 'EXAMPLE' -and $text) { $examples += $text }
      elseif ($curTag -eq 'CATEGORY' -and $text) { $category = $text }
      elseif ($curTag -eq 'LINK' -and $text) {
        $url = $null
        $firstNonEmpty = $null
        foreach ($ln in ($text -split "`n")) {
          $tln = $ln.Trim()
          if (-not $tln) { continue }
          if ($tln -match '^(?i)(https?://|www\.|mailto:|file:)') { $url = $tln; break }
          if (-not $firstNonEmpty) { $firstNonEmpty = $tln }
        }
        if ($url) { $category = $url }
        elseif ($firstNonEmpty) { $category = $firstNonEmpty }
      }
      break
    }
    if ($line -match '^\s*\.(\w+)\b') {
      # flush previous tag
      $text = ($buf -join " `n").Trim()
      if ($curTag -eq 'SYNOPSIS' -and $text) { $syn = $text }
      elseif ($curTag -eq 'DESCRIPTION' -and $text) { $desc = $text }
      elseif ($curTag -eq 'EXAMPLE' -and $text) { $examples += $text }
      elseif ($curTag -eq 'CATEGORY' -and $text) { $category = $text }
      elseif ($curTag -eq 'LINK' -and $text) {
        $url = $null
        $firstNonEmpty = $null
        foreach ($ln in ($text -split "`n")) {
          $tln = $ln.Trim()
          if (-not $tln) { continue }
          if ($tln -match '^(?i)(https?://|www\.|mailto:|file:)') { $url = $tln; break }
          if (-not $firstNonEmpty) { $firstNonEmpty = $tln }
        }
        if ($url) { $category = $url }
        elseif ($firstNonEmpty) { $category = $firstNonEmpty }
      }
      # start new tag
      $curTag = $Matches[1].ToUpperInvariant()
      $buf = @()
      $content = ($line -replace '^\s*\.\w+\s*', '').Trim()
      if ($content) { $buf += $content }
      continue
    }
    if ($curTag) { $buf += $line }
  }
  if (-not $syn) { $syn = $desc }
  if (-not $desc) { $desc = $syn }
  [pscustomobject]@{ Synopsis = $syn; Description = $desc; Examples = $examples; Category = $category }
}

function Get-ParameterMeta([System.Management.Automation.CommandInfo]$cmd) {
  $arr = @()
  if (-not $cmd -or -not $cmd.Parameters) { return @() }
  foreach ($kv in $cmd.Parameters.GetEnumerator()) {
    $p = $kv.Value
    $name = $p.Name
    $typeName = if ($p.ParameterType) { $p.ParameterType.Name } else { 'String' }
    $pos = [int]::MinValue
    $required = $false
    $attrs = @()
    try { $attrs = @($p.Attributes) } catch { $attrs = @() }
    $paramAttr = $null
    foreach ($a in $attrs) {
      if ($a -is [System.Management.Automation.ParameterAttribute]) { $paramAttr = $a; break }
    }
    if ($paramAttr) {
      try { if ($null -ne $paramAttr.Position) { $pos = [int]$paramAttr.Position } } catch { $pos = [int]::MinValue }
      try { $required = [bool]$paramAttr.Mandatory } catch { $required = $false }
    }
    $arr += [pscustomobject]@{ Name = $name; Required = $required; Type = $typeName; Position = $pos }
  }
  return ($arr | Sort-Object Position, Name)
}

function Filter-UserParameters([object[]]$params) {
  $common = @(
    'Verbose', 'Debug', 'ErrorAction', 'ErrorVariable', 'WarningAction', 'WarningVariable',
    'InformationAction', 'InformationVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable',
    'WhatIf', 'Confirm'
  )
  $params |
  Where-Object { $_ -and $_.PSObject -and $_.PSObject.Properties['Name'] } |
  Where-Object { $name = $_.Name; $name -and -not ($common -contains $name) }
}

function Build-Model([string]$functionsPath, [string]$CategoryRoot) {
  if (-not (Test-Path $functionsPath)) { throw "Root not found: $functionsPath" }
  $selfPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
  $all = Get-ChildItem -Path $functionsPath -Recurse -File -Filter *.ps1 -EA SilentlyContinue |
  Where-Object { $_.FullName -ne $selfPath -and $_.FullName -notmatch "\\\.git\\" }
  $funcFiles = @()
  foreach ($f in $all) {
    $fn = Detect-FunctionNameInFile $f
    if ($fn -and ($fn -eq [IO.Path]::GetFileNameWithoutExtension($f.Name))) { $funcFiles += $f }
  }
  # Import function files so their functions are callable
  if (-not $ScanSafe) {
    [void](Load-ModuleFromFiles -files $funcFiles)
  }

  # We'll resolve a repository root per file unless a root hint is provided
  $rootHint = $null
  if ($CategoryRoot) {
    try { $rootHint = [IO.Path]::GetFullPath($CategoryRoot) } catch { $rootHint = $CategoryRoot }
  }

  $items = @()
  # Clamp root fallback: compute an initial scan root to avoid climbing outside when resolving per-file
  $scanRoot = try { [IO.Path]::GetFullPath($functionsPath) } catch { $functionsPath }
  $funcSet = New-Object 'System.Collections.Generic.HashSet[string]'
  foreach ($f in $funcFiles) { [void]$funcSet.Add($f.FullName) }

  # Function commands
  foreach ($f in $funcFiles) {
    $fn = [IO.Path]::GetFileNameWithoutExtension($f.Name)
    $help = Parse-HelpFromFile -filePath $f.FullName
    $params = @()
    if (-not $ScanSafe) {
      $cmd = Get-Command -Name $fn -EA SilentlyContinue
      if ($cmd) { $params = Get-ParameterMeta -cmd $cmd; $params = @(Filter-UserParameters -params $params) }
    }
    $rootPerFile = if ($rootHint) { $rootHint } else { Resolve-CategoryRoot -path $f.FullName }
    # If Resolve-CategoryRoot returned a path outside of scan root, clamp it back
    try {
      $rp = [IO.Path]::GetFullPath($rootPerFile)
      if ($rp -and $scanRoot -and -not ($rp.ToLowerInvariant().StartsWith($scanRoot.ToLowerInvariant()))) {
        $rootPerFile = $scanRoot
      }
    }
    catch { $rootPerFile = $scanRoot }
  $cat = Get-CategoryForItem -cmdName $fn -filePath $f.FullName -root $rootPerFile -parsedCat $help.Category -DisableHelp:$ScanSafe
    $items += [pscustomobject]@{
      Name        = $fn
      Category    = $cat
      Synopsis    = $help.Synopsis
      Description = $help.Description
      Examples    = $help.Examples
      Parameters  = $params
      Invocation  = $(if ($ScanSafe) { $f.FullName } else { $fn })
      ScriptPath  = $f.FullName
    }
  }

  # Standalone scripts (not imported as functions)
  foreach ($f in $all) {
    if ($funcSet.Contains($f.FullName)) { continue }
    $cmd = $null
    if (-not $ScanSafe) { $cmd = Get-Command -Name $f.FullName -EA SilentlyContinue }
    $help = Parse-HelpFromFile -filePath $f.FullName
    $params = @()
    if ($cmd) { $params = Get-ParameterMeta -cmd $cmd; $params = @(Filter-UserParameters -params $params) }
    # Display concise base name for scripts to avoid overflow
    $display = [IO.Path]::GetFileNameWithoutExtension($f.Name)
    $rootPerFile = if ($rootHint) { $rootHint } else { Resolve-CategoryRoot -path $f.FullName }
    try {
      $rp = [IO.Path]::GetFullPath($rootPerFile)
      if ($rp -and $scanRoot -and -not ($rp.ToLowerInvariant().StartsWith($scanRoot.ToLowerInvariant()))) {
        $rootPerFile = $scanRoot
      }
    }
    catch { $rootPerFile = $scanRoot }
  $cat = Get-CategoryForItem -cmdName $f.FullName -filePath $f.FullName -root $rootPerFile -parsedCat $help.Category -DisableHelp:$ScanSafe
    $items += [pscustomobject]@{
      Name        = $display
      Category    = $cat
      Synopsis    = $help.Synopsis
      Description = $help.Description
      Examples    = $help.Examples
      Parameters  = $params
      Invocation  = $f.FullName
      ScriptPath  = $f.FullName
    }
  }

  # Normalize and group
  $norm = foreach ($it in $items) {
    [pscustomobject]@{
      Name        = $it.Name
      Category    = (Normalize-Category $it.Category)
      Synopsis    = $it.Synopsis
      Description = $it.Description
      Examples    = $it.Examples
      Parameters  = $it.Parameters
      Invocation  = $it.Invocation
      ScriptPath  = $it.ScriptPath
    }
  }
  $groups = $norm | Group-Object Category | Sort-Object Name
  $model = foreach ($g in $groups) { [pscustomobject]@{ Category = $g.Name; Commands = ($g.Group | Sort-Object Name) } }
  return $model
}

function Prompt-And-Run([pscustomobject]$cmdMeta) {
  [Console]::Clear()
  Write-Host ("Running: " + $cmdMeta.Name) -ForegroundColor $script:Theme.Info
  Write-Host ""

  # Order: required first (by position if any), then optional
  $paramListAll = @($cmdMeta.Parameters)
  $paramListAll = $paramListAll |
  Where-Object { $_ -and $_.PSObject -and $_.PSObject.Properties['Name'] } |
  ForEach-Object { $_ }
  # Ensure arrays even when only one element survives sorting
  $req = @(( $paramListAll | Where-Object { $_.Required } | Sort-Object Position ))
  $opt = @(( $paramListAll | Where-Object { -not $_.Required } | Sort-Object Name ))
  $ordered = @($req + $opt)

  $splat = @{}
  foreach ($p in $ordered) {
    $pTypeName = 'String'
    try {
      if ($null -ne $p) {
        if ($p.PSObject -and $p.PSObject.Properties['Type']) {
          $pTypeName = "$($p.Type)"
        }
        elseif ($p.PSObject -and $p.PSObject.Properties['ParameterType']) {
          $pTypeName = if ($p.ParameterType -and $p.ParameterType.Name) { $p.ParameterType.Name } else { "$($p.ParameterType)" }
        }
      }
    }
    catch { $pTypeName = 'String' }
    # Ensure we have a valid parameter name; skip otherwise
    $pNameVal = if ($p -and $p.PSObject.Properties['Name']) { "$($p.Name)" } else { $null }
    if (-not $pNameVal) { continue }
    $label = "{0} ({1}{2})" -f $pNameVal, $pTypeName, ($(if ($p.Required) { '; required' } else { '' }))
    if ($pTypeName -match '^(SwitchParameter|Switch)$') {
      $ans = Prompt-String "$label - toggle [y/N]" ''
      if ($ans -match '^(y|yes|true|1)$') { $splat[$pNameVal] = $true }
    }
    elseif ($pTypeName -match 'PSCredential') {
      $ans = Prompt-String "$label - press Enter to prompt for credential or leave blank to skip" ''
      if ($ans -ne '') {
        $cast = $null
        if (Try-Cast $ans $pTypeName ([ref]$cast)) { $splat[$pNameVal] = $cast }
      }
      else {
        $cred = Get-Credential -Message "Enter PSCredential for -$($p.Name) (Esc to cancel)"
        if ($cred) { $splat[$pNameVal] = $cred }
      }
    }
    else {
      $ans = Prompt-String $label ''
      if ($ans -ne '') {
        $cast = $null
        if (Try-Cast $ans $pTypeName ([ref]$cast)) { $splat[$pNameVal] = $cast }
        else { $splat[$pNameVal] = $ans }
      }
      elseif ($p.Required) {
        Write-Host "Parameter -$($p.Name) is required. Please enter a value." -ForegroundColor $script:Theme.Warn
        $ans = Prompt-String $label ''
        if ($ans -ne '') {
          $cast = $null
          if (Try-Cast $ans $pTypeName ([ref]$cast)) { $splat[$pNameVal] = $cast } else { $splat[$pNameVal] = $ans }
        }
      }
    }
  }

  Write-Host ""
  Write-Host "Command preview:" -ForegroundColor $script:Theme.Subtle
  $target = if ($cmdMeta.PSObject.Properties['Invocation']) { $cmdMeta.Invocation } else { $cmdMeta.Name }
  $preview = $target
  foreach ($k in $splat.Keys) {
    $v = $splat[$k]
    if ($v -is [switch] -or ($v -is [bool] -and $v -eq $true)) {
      $preview += " -$k"
    }
    elseif ($v -is [securestring]) {
      $preview += " -$k ********"
    }
    elseif ($v -is [string[]]) {
      $preview += " -$k `"$($v -join ',')`""
    }
    else {
      $preview += " -$k `"$v`""
    }
  }
  Write-Host "`n$preview`n" -ForegroundColor $script:Theme.Preview

  $confirm = Read-Host "Press Enter to run, or type 'n' to cancel"
  if ($confirm -match '^(n|no)$') { return }

  # Execute the command in a 'raw' way: relax strict mode and error preferences
  # so the called functions behave like when run standalone.
  Write-Host ""
  $prevEA = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  try {
    try { Set-StrictMode -Off } catch { }
    $target = if ($cmdMeta.PSObject.Properties['Invocation']) { $cmdMeta.Invocation } else { $cmdMeta.Name }
    & (Get-Command $target) @splat | Out-Host
  }
  catch {
    # Surface the original error without extra TUI formatting
    Write-Error $_
  }
  finally {
    $ErrorActionPreference = $prevEA
    try { Set-StrictMode -Version Latest } catch { }
  }
  Write-Host ""
  Write-Host "Press any key to return..." -ForegroundColor $script:Theme.Subtle
  [void](Read-Key)
}

function Run-TUI([object[]]$model) {
  # Cache model groups once
  $cats = @($model)
  $catIdx = 0
  $cmdIdx = 0
  $focus = 'cat' # cat | cmd | detail
  $scrollCat = 0
  $scrollCmd = 0
  $scrollDetail = 0
  $detailLines = @()

  # Track size and dirty state
  $lastW = -1; $lastH = -1
  $needFull = $true
  $dirtyCats = $true
  $dirtyCmds = $true
  $dirtyDetails = $true
  $prevCatIdx = -1
  $prevCmdIdx = -1

  while ($true) {
    $w = [Console]::WindowWidth
    $h = [Console]::WindowHeight
    if ($w -ne $lastW -or $h -ne $lastH) {
      $lastW = $w; $lastH = $h
      $needFull = $true
      $dirtyCats = $true; $dirtyCmds = $true; $dirtyDetails = $true
    }

  $header = " VTS Tools Browser "
    $footer = if ($ThemeStyle -eq 'minimal') {
      " Left/Right change pane  |  Up/Down select or scroll  |  Enter run  |  Ctrl+C exit "
    } else {
      " Left/Right: change pane  |  Up/Down: select or scroll  |  Enter: run  |  Ctrl+C: exit "
    }

    # Layout
    $paneTop = 1
    $paneHeight = $h - 2
    $leftW = [Math]::Max(18, [Math]::Min(30, [int]($w * 0.22)))
    $midW = [Math]::Max(24, [Math]::Min(42, [int]($w * 0.28)))
    $rightW = $w - $leftW - $midW - 4

  if ($needFull) {
      [Console]::Clear()
      # Header and footer
      $hdrBg = $(if ($ThemeStyle -eq 'minimal') { $script:Theme.Background } else { $script:Theme.HeaderBg })
      $hdrFg = $(if ($ThemeStyle -eq 'minimal') { $script:Theme.Subtle } else { $script:Theme.HeaderFg })
      $ftrBg = $(if ($ThemeStyle -eq 'minimal') { $script:Theme.Background } else { $script:Theme.FooterBg })
      $ftrFg = $(if ($ThemeStyle -eq 'minimal') { $script:Theme.Subtle } else { $script:Theme.FooterFg })
  Draw-Line 0 0 $w ($hdrBg) ' '
  # Center header
  $hpad = [Math]::Max(0, [int](($w - $header.Length) / 2))
  Write-At 0 $hpad (TruncatePad $header ($w - $hpad)) ($hdrFg) $hdrBg -NoNewline
      Draw-Line ($h - 1) 0 $w ($ftrBg) ' '
  # Center footer
  $fpad = [Math]::Max(0, [int](($w - $footer.Length) / 2))
  Write-At ($h - 1) $fpad (TruncatePad $footer ($w - $fpad)) ($ftrFg) $ftrBg -NoNewline
  # Frames drawn below based on active focus color
      # Fill pane backgrounds
      Fill-Rect ($paneTop + 1) 1 ($paneHeight - 2) ($leftW - 2) $script:Theme.Background
      Fill-Rect ($paneTop + 1) ($leftW + 2) ($paneHeight - 2) ($midW - 2) $script:Theme.Background
      Fill-Rect ($paneTop + 1) ($leftW + $midW + 3) ($paneHeight - 2) ($rightW - 2) $script:Theme.Background
      # Minimal style: draw subtle vertical separators so panes are visually divided
      if ($ThemeStyle -eq 'minimal') {
        $sepChar = if ($NoUnicode) { '|' } else { "$([char]0x2502)" } # '│'
        Draw-VertLine ($paneTop + 1) ($leftW - 1) ($paneHeight - 2) ($script:Theme.Border) $sepChar
        Draw-VertLine ($paneTop + 1) ($leftW + $midW) ($paneHeight - 2) ($script:Theme.Border) $sepChar
      }
    }

    # Draw frames (or redraw when focus changes)
    if (-not (Get-Variable -Name prevFocus -Scope Script -EA SilentlyContinue)) { $script:prevFocus = $null }
    $framesDirty = $needFull -or ($script:prevFocus -ne $focus)
    if ($framesDirty) {
      $script:prevFocus = $focus
      $colCat = if ($focus -eq 'cat') { $script:Theme.Accent } else { $script:Theme.Border }
      $colCmd = if ($focus -eq 'cmd') { $script:Theme.Accent } else { $script:Theme.Border }
      $colDet = if ($focus -eq 'detail') { $script:Theme.Accent } else { $script:Theme.Border }
      Box $paneTop 0 $paneHeight $leftW "Categories" $colCat
      Box $paneTop ($leftW + 1) $paneHeight $midW "Commands" $colCmd
      Box $paneTop ($leftW + $midW + 2) $paneHeight $rightW "Details" $colDet
    }

    if (-not $cats) {
      Write-At ([int]($paneTop + 2)) 2 "No commands found." ($script:Theme.Warn) $script:Theme.Background
      Write-At ($h - 2) 2 "Press any key to refresh." ($script:Theme.Subtle) $script:Theme.Background
      [void](Read-Key)
      $needFull = $true; $dirtyCats = $true; $dirtyCmds = $true; $dirtyDetails = $true
      continue
    }

    if ($catIdx -ge $cats.Count) { $catIdx = 0 }
    $selectedCat = $cats[$catIdx]
    $cmds = @()
    if ($selectedCat -and ($selectedCat.PSObject.Properties['Commands'])) { $cmds = @($selectedCat.Commands) }
    if ($cmdIdx -ge $cmds.Count) { $cmdIdx = 0 }

    # Determine dirty by selection change
    if ($prevCatIdx -ne $catIdx) { $dirtyCats = $true; $dirtyCmds = $true; $dirtyDetails = $true; $scrollDetail = 0; $prevCatIdx = $catIdx }
    if ($prevCmdIdx -ne $cmdIdx) { $dirtyCmds = $true; $dirtyDetails = $true; $scrollDetail = 0; $prevCmdIdx = $cmdIdx }

    # Render categories when dirty
    if ($dirtyCats -or $needFull) {
      $listTop = $paneTop + 1
      $listHeight = $paneHeight - 2
      if ($catIdx -lt $scrollCat) { $scrollCat = $catIdx }
      if ($catIdx -ge ($scrollCat + $listHeight)) { $scrollCat = $catIdx - $listHeight + 1 }

      for ($i = 0; $i -lt $listHeight; $i++) {
        $idx = $scrollCat + $i
        $y = $listTop + $i
        $line = ' ' * ($leftW - 2)
        if ($idx -lt $cats.Count) {
          $catItem = $cats[$idx]
          $name = if ($catItem -and ($catItem.PSObject.Properties['Category'])) { $catItem.Category } else { 'Uncategorized' }
          $count = if ($catItem -and ($catItem.PSObject.Properties['Commands'])) { @($catItem.Commands).Count } else { 0 }
          $txt = ("{0} ({1})" -f $name, $count)
          $bullet = ($script:Symbols.Bullet.Trim())
          $txt2 = ("{0} {1}" -f $bullet, $txt)
          $line = TruncatePad $txt2 ($leftW - 2)
        }
        $isSel = ($idx -eq $catIdx)
        if ($isSel -and $focus -eq 'cat') { $fg = $script:Theme.CatSelFg; $bg = $script:Theme.CatSelBg }
        else { $fg = $(if ($ThemeStyle -eq 'minimal') { $script:Theme.Subtle } else { $script:Theme.Foreground }); $bg = $script:Theme.Background }
        [Console]::SetCursorPosition(1, $y); [Console]::ForegroundColor = $fg; [Console]::BackgroundColor = $bg; [Console]::Write($line); [Console]::ResetColor()
      }
      # Draw scroll indicators (far-right column) when overflow exists
      $xIndCat = ($leftW - 2)
      if ($scrollCat -gt 0 -and $cats.Count -gt 0) {
        $idxTop = $scrollCat
        $isSelTop = ($idxTop -eq $catIdx -and $focus -eq 'cat')
        $bgTop = if ($isSelTop) { $script:Theme.CatSelBg } else { $script:Theme.Background }
        [Console]::SetCursorPosition($xIndCat, $listTop)
        [Console]::BackgroundColor = $bgTop; [Console]::ForegroundColor = $script:Theme.Accent; [Console]::Write($script:Symbols.Up); [Console]::ResetColor()
      }
      if ( ($scrollCat + $listHeight) -lt $cats.Count ) {
        $idxBot = [Math]::Min($cats.Count - 1, $scrollCat + $listHeight - 1)
        $isSelBot = ($idxBot -eq $catIdx -and $focus -eq 'cat')
        $bgBot = if ($isSelBot) { $script:Theme.CatSelBg } else { $script:Theme.Background }
        [Console]::SetCursorPosition($xIndCat, ($listTop + $listHeight - 1))
        [Console]::BackgroundColor = $bgBot; [Console]::ForegroundColor = $script:Theme.Accent; [Console]::Write($script:Symbols.Down); [Console]::ResetColor()
      }
      $dirtyCats = $false
    }

    # Render commands when dirty
    if ($dirtyCmds -or $needFull) {
      $cmdTop = $paneTop + 1
      $cmdHeight = $paneHeight - 2
      if ($cmdIdx -lt $scrollCmd) { $scrollCmd = $cmdIdx }
      if ($cmdIdx -ge ($scrollCmd + $cmdHeight)) { $scrollCmd = $cmdIdx - $cmdHeight + 1 }

      for ($i = 0; $i -lt $cmdHeight; $i++) {
        $idx = $scrollCmd + $i
        $y = $cmdTop + $i
        $line = ' ' * ($midW - 2)
        if ($idx -lt $cmds.Count) {
          $txt = $cmds[$idx].Name
          $bullet = ($script:Symbols.Bullet.Trim())
          $txt2 = ("{0} {1}" -f $bullet, $txt)
          $line = TruncatePad $txt2 ($midW - 2)
        }
        $isSel = ($idx -eq $cmdIdx)
        if ($isSel -and $focus -eq 'cmd') { $fg = $script:Theme.CmdSelFg; $bg = $script:Theme.CmdSelBg }
        else { $fg = $(if ($ThemeStyle -eq 'minimal') { $script:Theme.Subtle } else { $script:Theme.Foreground }); $bg = $script:Theme.Background }
        [Console]::SetCursorPosition($leftW + 2, $y); [Console]::ForegroundColor = $fg; [Console]::BackgroundColor = $bg; [Console]::Write($line); [Console]::ResetColor()
      }
      # Draw scroll indicators (far-right column) when overflow exists
      $xIndCmd = ($leftW + $midW - 1)
      if ($scrollCmd -gt 0 -and $cmds.Count -gt 0) {
        $idxTop = $scrollCmd
        $isSelTop = ($idxTop -eq $cmdIdx -and $focus -eq 'cmd')
        $bgTop = if ($isSelTop) { $script:Theme.CmdSelBg } else { $script:Theme.Background }
        [Console]::SetCursorPosition($xIndCmd, $cmdTop)
        [Console]::BackgroundColor = $bgTop; [Console]::ForegroundColor = $script:Theme.Accent; [Console]::Write($script:Symbols.Up); [Console]::ResetColor()
      }
      if ( ($scrollCmd + $cmdHeight) -lt $cmds.Count ) {
        $idxBot = [Math]::Min($cmds.Count - 1, $scrollCmd + $cmdHeight - 1)
        $isSelBot = ($idxBot -eq $cmdIdx -and $focus -eq 'cmd')
        $bgBot = if ($isSelBot) { $script:Theme.CmdSelBg } else { $script:Theme.Background }
        [Console]::SetCursorPosition($xIndCmd, ($cmdTop + $cmdHeight - 1))
        [Console]::BackgroundColor = $bgBot; [Console]::ForegroundColor = $script:Theme.Accent; [Console]::Write($script:Symbols.Down); [Console]::ResetColor()
      }
      $dirtyCmds = $false
    }

    # Render details when dirty
    if ($dirtyDetails -or $needFull) {
      $detail = if ($cmds.Count) { $cmds[$cmdIdx] } else { $null }
      $dTop = $paneTop + 1
      $dHeight = $paneHeight - 2
      $wrapWidth = $rightW - 3
      # Clear details area with themed background
      for ($yy = 0; $yy -lt $dHeight; $yy++) {
        [Console]::SetCursorPosition($leftW + $midW + 3, $dTop + $yy)
        [Console]::BackgroundColor = $script:Theme.Background
        [Console]::Write((' ' * $wrapWidth))
        [Console]::ResetColor()
      }
      $lineY = $dTop
      if ($detail) {
        $lines = @()
        # Header
        $lines += "Name: $($detail.Name)"
        $lines += "Category: $($detail.Category)"
        # Synopsis
        if ($detail.Synopsis) {
          $lines += ''
          $lines += 'Synopsis:'
          $lines += @(Wrap-Block $detail.Synopsis $wrapWidth)
        }
        # Description
        if ($detail.Description) {
          $lines += ''
          $lines += 'Description:'
          $lines += @(Wrap-Block $detail.Description $wrapWidth)
        }
        # Parameters
        $paramList = @($detail.Parameters)
        if ($paramList.Count) {
          $lines += ''
          $lines += 'Parameters:'
          foreach ($p in $paramList | Sort-Object { -not $_.Required }, Position, Name) {
            $req = if ($p.Required) { 'required' } else { 'optional' }
            $pline = "$($script:Symbols.Bullet)$($p.Name) <$($p.Type)> ($req)"
            $lines += @(Wrap-Bullet $pline $wrapWidth ' ' '   ')
          }
        }
        # Examples
        $exList = @($detail.Examples)
        if ($exList.Count) {
          $lines += ''
          $lines += 'Examples:'
          $firstEx = $true
          foreach ($ex in $exList | Select-Object -First 3) {
            if (-not $firstEx) {
              # Ensure exactly one blank line between examples: trim trailing blanks, then add one
              while ($lines.Count -gt 0 -and $lines[-1] -eq '') { $lines = $lines[0..($lines.Count - 2)] }
              $lines += ''
            }
            $firstEx = $false
            $exText = ("$ex" -replace "\r\n", "`n") -replace "\r", "`n"
            $exLines = @($exText -split "`n")
            $isFirstLine = $true
      foreach ($el in $exLines) {
              # Keep empty lines inside example
              if ($el -eq $null -or $el.Length -eq 0) { $lines += ''; continue }
              $t = $el.TrimEnd()
              if ($isFirstLine) {
                # Command line: bullet and wrap nicely
                $lines += @(Wrap-Bullet ("$($script:Symbols.Bullet)$t") $wrapWidth ' ' '   ')
                $isFirstLine = $false
        # Add an extra blank line between the command and its output
        $lines += ''
              }
              else {
                # Output line(s): preserve spacing and wrap hard at width with slight indent
                $lit = "  $t"
                $lines += @(Wrap-LineLiteral $lit $wrapWidth)
              }
            }
          }
        }
        # Update cached lines and clamp scroll
        $detailLines = @($lines)
        $maxScroll = [Math]::Max(0, $detailLines.Count - $dHeight)
        if ($scrollDetail -gt $maxScroll) { $scrollDetail = $maxScroll }
        if ($scrollDetail -lt 0) { $scrollDetail = 0 }

        # Render a visible slice with section highlighting
        $rendered = 0
        $headings = @('Synopsis:', 'Description:', 'Parameters:', 'Examples:')
        for ($i = $scrollDetail; $i -lt $detailLines.Count -and $rendered -lt $dHeight; $i++) {
          $l = $detailLines[$i]
          $xBase = ($leftW + $midW + 3)
          if ($l -like 'Name:*' -or $l -like 'Category:*') {
            $m = [regex]::Match($l, '^(Name|Category):\s*(.*)$')
            if ($m.Success) {
              $label = $m.Groups[1].Value + ': '
              $val = $m.Groups[2].Value
              $valTrunc = TruncatePad $val ([Math]::Max(0, $wrapWidth - $label.Length))
              Write-At $lineY $xBase $label $script:Theme.SectionHdr $script:Theme.Background -NoNewline
              Write-At $lineY ($xBase + $label.Length) $valTrunc $script:Theme.Text $script:Theme.Background
            }
            else {
              Write-At $lineY $xBase (TruncatePad $l $wrapWidth) $script:Theme.Text $script:Theme.Background
            }
          }
          elseif ($headings -contains $l) {
            $title = if ($ThemeStyle -eq 'minimal') { $l } else { ("{0} {1}" -f $script:Symbols.Sparkle, $l) }
            Write-At $lineY $xBase (TruncatePad $title $wrapWidth) $script:Theme.SectionHdr $script:Theme.Background
          }
          else {
            Write-At $lineY $xBase (TruncatePad $l $wrapWidth) $script:Theme.Foreground $script:Theme.Background
          }
          $lineY++
          $rendered++
        }
        # Draw scroll indicators at far-right column if overflow
        $xInd = ($leftW + $midW + 3) + ($wrapWidth - 1)
        if ($scrollDetail -gt 0) {
          # Up indicator at first visible line
          Write-At $dTop $xInd $script:Symbols.Up ($script:Theme.Accent) $script:Theme.Background
        }
        if (($scrollDetail + $dHeight) -lt $detailLines.Count) {
          # Down indicator at last visible line
          Write-At ($dTop + $dHeight - 1) $xInd $script:Symbols.Down ($script:Theme.Accent) $script:Theme.Background
        }
      }
      else {
        Write-At $dTop ($leftW + $midW + 3) "No commands in this category." ($script:Theme.Warn) $script:Theme.Background
      }
      $dirtyDetails = $false
    }

    $needFull = $false

    # Key handling
    $k = Read-Key
    switch ($k.Key) {
      'Enter' {
        $detail = $null
        $selectedCat = if ($cats.Count) { $cats[$catIdx] } else { $null }
        if ($selectedCat -and ($selectedCat.PSObject.Properties['Commands'])) { $cmds = @($selectedCat.Commands) } else { $cmds = @() }
        if ($cmds.Count) { $detail = $cmds[$cmdIdx] }
        if ($detail) { Prompt-And-Run $detail }
        $needFull = $true; $dirtyCats = $true; $dirtyCmds = $true; $dirtyDetails = $true
      }
      'UpArrow' {
        switch ($focus) {
          'cat' { if ($catIdx -gt 0) { $catIdx--; $cmdIdx = 0 } }
          'cmd' { if ($cmdIdx -gt 0) { $cmdIdx-- } }
          'detail' { if ($scrollDetail -gt 0) { $scrollDetail--; $dirtyDetails = $true } }
        }
      }
      'DownArrow' {
        switch ($focus) {
          'cat' { if ($catIdx -lt ($cats.Count - 1)) { $catIdx++; $cmdIdx = 0 } }
          'cmd' { if ($cmdIdx -lt ($cmds.Count - 1)) { $cmdIdx++ } }
          'detail' {
            $dHeight = $paneHeight - 2
            $maxScroll = [Math]::Max(0, @($detailLines).Count - $dHeight)
            if ($scrollDetail -lt $maxScroll) { $scrollDetail++; $dirtyDetails = $true }
          }
        }
      }
      'LeftArrow' {
        if ($focus -eq 'cmd') { $focus = 'cat' }
        elseif ($focus -eq 'detail') { $focus = 'cmd' }
        $dirtyCats = $true; $dirtyCmds = $true; $dirtyDetails = $true
      }
      'RightArrow' {
        if ($focus -eq 'cat') { $focus = 'cmd' }
        elseif ($focus -eq 'cmd') { $focus = 'detail' }
        $dirtyCats = $true; $dirtyCmds = $true; $dirtyDetails = $true
      }
      default { }
    }
  }
}

try {
  Write-Verbose "Starting Ensure-Console"
  Ensure-Console
  Write-Verbose "Ensure-Console OK"
  # Normalize/resolve FunctionsPath if empty or invalid
  if (-not $FunctionsPath -or [string]::IsNullOrWhiteSpace($FunctionsPath)) {
    $FunctionsPath = Resolve-FunctionsPath -inputPath $FunctionsPath
    Write-Host "[DEBUG] Auto-resolved FunctionsPath to '$FunctionsPath'" -ForegroundColor $script:Theme.Subtle
  }
  elseif (-not (Test-Path -LiteralPath $FunctionsPath)) {
    $FunctionsPath = Resolve-FunctionsPath -inputPath $FunctionsPath
    Write-Host "[DEBUG] Adjusted FunctionsPath to '$FunctionsPath'" -ForegroundColor $script:Theme.Subtle
  }

  # Early bootstrap: if FunctionsPath doesn't contain any .ps1 besides this file, and a repo URL is configured, fetch and run from repo
  $shouldBootstrap = $false
  try {
    if (-not (Test-Path $FunctionsPath)) { $shouldBootstrap = $true }
    else {
      $selfPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
      $hasAny = Get-ChildItem -Path $FunctionsPath -Recurse -File -Filter *.ps1 -ErrorAction SilentlyContinue | Where-Object { $_.FullName -ne $selfPath -and $_.FullName -notmatch "\\\.git\\" } | Select-Object -First 1
      if (-not $hasAny) { $shouldBootstrap = $true }
    }
  }
  catch { $shouldBootstrap = $true }

  if ($shouldBootstrap -and $BootstrapFromRepoUrl) {
    # If a repo URL is explicitly provided, always bootstrap regardless of local content
    Write-Verbose ("Bootstrapping from {0} (branch={1}, start={2})" -f $BootstrapFromRepoUrl, $Branch, $StartScriptRelPath)
    Install-And-RunFromRepo -repoUrl $BootstrapFromRepoUrl -branch $Branch -startRelPath $StartScriptRelPath -installBase $InstallBase -reinstall:$Reinstall
    return
  }
  elseif ($BootstrapFromRepoUrl) {
    Write-Verbose ("Bootstrapping (forced) from {0} (branch={1}, start={2})" -f $BootstrapFromRepoUrl, $Branch, $StartScriptRelPath)
    Install-And-RunFromRepo -repoUrl $BootstrapFromRepoUrl -branch $Branch -startRelPath $StartScriptRelPath -installBase $InstallBase -reinstall:$Reinstall
    return
  }
  elseif ($shouldBootstrap) {
    Write-Verbose "No local commands and no repo URL provided; continuing without bootstrap"
  }

  Write-Verbose ("Building model from {0}" -f $FunctionsPath)
  $model = Build-Model -functionsPath $FunctionsPath
  $catCount = if ($null -eq $model) { 0 } else { @($model).Count }
  Write-Verbose ("Model built: {0} categories" -f $catCount)
  if ($ValidateOnly) {
    $totalCmds = 0
    foreach ($g in @($model)) { $totalCmds += @($g.Commands).Count }
    Write-Host "Validation summary:" -ForegroundColor $script:Theme.Info
    Write-Host (" - Categories: {0}" -f $catCount)
    Write-Host (" - Total commands: {0}" -f $totalCmds)
    foreach ($g in @($model)) {
      $name = if ($g.PSObject.Properties['Category']) { "$($g.Category)" } else { 'Uncategorized' }
      $cnt = if ($g.PSObject.Properties['Commands']) { @($g.Commands).Count } else { 0 }
      Write-Host ("   * {0} ({1})" -f $name, $cnt) -ForegroundColor $script:Theme.Subtle
    }
    return
  }
  Write-Verbose "Launching TUI"
  Run-TUI -model $model
}
catch {
  Write-Host "Error: $($_.Exception.Message)" -ForegroundColor $script:Theme.Error
  if ($_.InvocationInfo.PositionMessage) {
    Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor $script:Theme.Subtle
  }
  exit 1
}