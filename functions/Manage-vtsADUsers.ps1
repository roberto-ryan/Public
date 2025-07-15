function Manage-vtsADUsers {
  <#
  .SYNOPSIS
      Processes a list of Active Directory group memberships by comparing a target list with a source list.
  
  .DESCRIPTION
      This function takes a target group membership list and a source group membership list,
      then identifies items that are in both lists, only in the target list, or only in the source list.
      It returns a custom object containing these categorized memberships.
  
  .PARAMETER targetMembership
      The collection of group memberships considered as the target for comparison.
  
  .PARAMETER sourceMembership
      The collection of group memberships considered as the source for comparison.
  
  .OUTPUTS
      [PSCustomObject] with the following properties:
      - InBoth: Memberships that exist in both target and source lists
      - OnlyInTarget: Memberships that exist only in the target list
      - OnlyInSource: Memberships that exist only in the source list
  
  .EXAMPLE
      $targetMembers = Get-ADGroupMember -Identity "TargetGroup"
      $sourceMembers = Get-ADGroupMember -Identity "SourceGroup"
      $result = Process-GroupMembership -targetMembership $targetMembers -sourceMembership $sourceMembers
      $result.InBoth | ForEach-Object { Write-Host "Member in both groups: $($_)" }
  
  .NOTES
      This function uses Compare-Object to efficiently identify differences between the two lists.
  
  .LINK
      Active Directory
      #>
  # Load required assemblies for Windows Forms
  Add-Type -AssemblyName System.Windows.Forms, System.Drawing
  
  # Setup logging function
  function Write-Log {
      param (
          [string]$Message,
          [string]$Level = "INFO"
      )
      $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
      $logMessage = "[$timestamp] [$Level] $Message"
      Write-Host $logMessage
      Add-Content -Path "$env:TEMP\ADUserManager.log" -Value $logMessage
  }
  
  Write-Host "`n========== AD User Manager ==========`n" -ForegroundColor Cyan
  Write-Host "This tool helps you manage Active Directory users." -ForegroundColor Cyan
  Write-Host "First, you'll select one or more OUs to work with." -ForegroundColor Cyan
  Write-Host "Then you can view and manage users in those OUs." -ForegroundColor Cyan
  Write-Host "======================================`n" -ForegroundColor Cyan
  
  Write-Log "Application started"
  
  # Fetch OUs and display in Out-GridView for selection
  try {
      Write-Host "Fetching organizational units from Active Directory..." -ForegroundColor Yellow
      $OUs = Get-ADOrganizationalUnit -Filter * -Properties Name,DistinguishedName | 
             Select-Object Name,DistinguishedName
      
      Write-Host "Select one or more OUs in the grid view window and click OK." -ForegroundColor Green
      $selectedOUs = $OUs | Out-GridView -Title "Select Organizational Units (Select multiple and click OK)" -OutputMode Multiple
      
      if ($null -eq $selectedOUs -or $selectedOUs.Count -eq 0) {
          Write-Host "No OUs selected. Exiting." -ForegroundColor Red
          Write-Log "No OUs selected, application terminated" -Level "WARN"
          exit
      }
      
      Write-Log "Selected $($selectedOUs.Count) OUs"
  } catch {
      Write-Host "Error fetching OUs: $_" -ForegroundColor Red
      Write-Log "Error fetching OUs: $_" -Level "ERROR"
      exit
  }
  
  # Fetch users from selected OUs
  try {
      Write-Host "Fetching users from selected OUs..." -ForegroundColor Yellow
      $users = @()
      foreach ($OU in $selectedOUs) {
          Write-Host "  Loading users from OU: $($OU.Name)" -ForegroundColor Gray
          # Get users with properties that help determine activity status
          $OUusers = Get-ADUser -Filter * -SearchBase $OU.DistinguishedName -Properties Name,SamAccountName,Enabled,LastLogonDate,PasswordLastSet,PasswordExpired,LockedOut,Description | 
                     Select-Object Name,SamAccountName,Enabled,LastLogonDate,PasswordLastSet,PasswordExpired,LockedOut,Description,DistinguishedName
          $users += $OUusers
      }
      
      if ($users.Count -eq 0) {
          Write-Host "No users found in selected OUs. Exiting." -ForegroundColor Red
          Write-Log "No users found in selected OUs" -Level "WARN"
          exit
      }
      
      Write-Log "Fetched $($users.Count) users from selected OUs"
  } catch {
      Write-Host "Error fetching users: $_" -ForegroundColor Red
      Write-Log "Error fetching users: $_" -Level "ERROR"
      exit
  }
  
  # Create the Windows Form
  $form = New-Object System.Windows.Forms.Form
  $form.Text = "AD User Manager"
  $form.Size = [System.Drawing.Size]::new(900, 600)
  $form.StartPosition = "CenterScreen"
  
  # Create the DataGridView to display users
  $grid = New-Object System.Windows.Forms.DataGridView
  $grid.Dock = "Fill"
  $grid.AutoSizeColumnsMode = "Fill"
  $grid.SelectionMode = "FullRowSelect"
  $grid.MultiSelect = $false
  $grid.AllowUserToAddRows = $false
  $grid.AllowUserToDeleteRows = $false
  $grid.ReadOnly = $true
  $form.Controls.Add($grid)
  
  # Create button panel
  $buttonPanel = New-Object System.Windows.Forms.Panel
  $buttonPanel.Dock = "Bottom"
  $buttonPanel.Height = 50
  $form.Controls.Add($buttonPanel)
  
  # Create the Refresh button
  $refreshButton = New-Object System.Windows.Forms.Button
  $refreshButton.Text = "Refresh Users"
  $refreshButton.Location = [System.Drawing.Point]::new(10, 10)
  $refreshButton.Size = [System.Drawing.Size]::new(120, 30)
  $buttonPanel.Controls.Add($refreshButton)
  
  # Create Enable User button
  $enableButton = New-Object System.Windows.Forms.Button
  $enableButton.Text = "Enable User"
  $enableButton.Location = [System.Drawing.Point]::new(140, 10)
  $enableButton.Size = [System.Drawing.Size]::new(120, 30)
  $buttonPanel.Controls.Add($enableButton)
  
  # Create Disable User button
  $disableButton = New-Object System.Windows.Forms.Button
  $disableButton.Text = "Disable User"
  $disableButton.Location = [System.Drawing.Point]::new(270, 10)
  $disableButton.Size = [System.Drawing.Size]::new(120, 30)
  $buttonPanel.Controls.Add($disableButton)
  
  # Create status bar for messages
  $statusBar = New-Object System.Windows.Forms.StatusStrip
  $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
  $statusLabel.Text = "Ready"
  $statusBar.Items.Add($statusLabel)
  $form.Controls.Add($statusBar)
  
  # Function to populate the grid with data
  function Populate-Grid {
      param ($data)
      
      # Create a DataTable for better compatibility with DataGridView
      $dataTable = New-Object System.Data.DataTable
      
      # Add columns to the DataTable
      "Name", "SamAccountName", "Enabled", "LastLogonDate", "PasswordLastSet", 
      "PasswordExpired", "LockedOut", "Description" | ForEach-Object {
          $dataTable.Columns.Add($_) | Out-Null
      }
      
      # Add rows to the DataTable
      foreach ($user in ($data | Sort-Object Enabled -Descending)) {
          $row = $dataTable.NewRow()
          $row["Name"] = $user.Name
          $row["SamAccountName"] = $user.SamAccountName
          $row["Enabled"] = $user.Enabled
          $row["LastLogonDate"] = $user.LastLogonDate
          $row["PasswordLastSet"] = $user.PasswordLastSet
          $row["PasswordExpired"] = $user.PasswordExpired
          $row["LockedOut"] = $user.LockedOut
          $row["Description"] = $user.Description
          $dataTable.Rows.Add($row)
      }
      
      # Set the DataSource to the DataTable
      $grid.DataSource = $dataTable
      
      Write-Log "Grid populated with $($data.Count) users"
  }
  
  # Function to refresh data from Active Directory
  function Refresh-Users {
      try {
          Write-Host "Refreshing user data..." -ForegroundColor Yellow
          $refreshedUsers = @()
          
          foreach ($OU in $selectedOUs) {
              Write-Host "  Refreshing users from OU: $($OU.Name)" -ForegroundColor Gray
              $OUusers = Get-ADUser -Filter * -SearchBase $OU.DistinguishedName -Properties Name,SamAccountName,Enabled,LastLogonDate,PasswordLastSet,PasswordExpired,LockedOut,Description | 
                         Select-Object Name,SamAccountName,Enabled,LastLogonDate,PasswordLastSet,PasswordExpired,LockedOut,Description,DistinguishedName
              $refreshedUsers += $OUusers
          }
          
          Populate-Grid -data $refreshedUsers
          $statusLabel.Text = "Users refreshed at $(Get-Date -Format 'HH:mm:ss')"
          Write-Log "Users refreshed" -Level "INFO"
      } catch {
          $statusLabel.Text = "Error refreshing users"
          [System.Windows.Forms.MessageBox]::Show("Error refreshing users: $_", "Error", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
          Write-Log "Error refreshing users: $_" -Level "ERROR"
      }
  }
  
  # Function to enable selected user
  function Enable-SelectedUser {
      if ($grid.SelectedRows.Count -eq 0) {
          $statusLabel.Text = "No user selected"
          [System.Windows.Forms.MessageBox]::Show("Please select a user to enable.", "No Selection", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
          return
      }
      
      $selectedRow = $grid.SelectedRows[0]
      $userName = $selectedRow.Cells["SamAccountName"].Value
      $userEnabled = $selectedRow.Cells["Enabled"].Value
      
      # Convert string representation to boolean if needed
      if ($userEnabled -is [string]) {
          $userEnabled = [System.Boolean]::Parse($userEnabled)
      }
      
      if ($userEnabled -eq $true) {
          $statusLabel.Text = "User already enabled"
          [System.Windows.Forms.MessageBox]::Show("$userName is already enabled.", "Info", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
          return
      }
      
      try {
          $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to enable $userName?", "Confirm", 
              [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
          
          if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
              Enable-ADAccount -Identity $userName
              Write-Log "User $userName enabled" -Level "INFO"
              $statusLabel.Text = "User $userName enabled successfully"
              [System.Windows.Forms.MessageBox]::Show("$userName has been enabled.", "Success", 
                  [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
              Refresh-Users
          }
      } catch {
          $statusLabel.Text = "Error enabling user"
          [System.Windows.Forms.MessageBox]::Show("Error enabling $userName`: $_", "Error", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
          Write-Log "Error enabling user $userName`: $_" -Level "ERROR"
      }
  }
  
  # Function to disable selected user
  function Disable-SelectedUser {
      if ($grid.SelectedRows.Count -eq 0) {
          $statusLabel.Text = "No user selected"
          [System.Windows.Forms.MessageBox]::Show("Please select a user to disable.", "No Selection", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
          return
      }
      
      $selectedRow = $grid.SelectedRows[0]
      $userName = $selectedRow.Cells["SamAccountName"].Value
      $userEnabled = $selectedRow.Cells["Enabled"].Value
      
      # Convert string representation to boolean if needed
      if ($userEnabled -is [string]) {
          $userEnabled = [System.Boolean]::Parse($userEnabled)
      }
      
      if ($userEnabled -eq $false) {
          $statusLabel.Text = "User already disabled"
          [System.Windows.Forms.MessageBox]::Show("$userName is already disabled.", "Info", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
          return
      }
      
      try {
          $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to disable $userName?", "Confirm", 
              [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
          
          if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
              Disable-ADAccount -Identity $userName
              Write-Log "User $userName disabled" -Level "INFO"
              $statusLabel.Text = "User $userName disabled successfully"
              [System.Windows.Forms.MessageBox]::Show("$userName has been disabled.", "Success", 
                  [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
              Refresh-Users
          }
      } catch {
          $statusLabel.Text = "Error disabling user"
          [System.Windows.Forms.MessageBox]::Show("Error disabling $userName`: $_", "Error", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
          Write-Log "Error disabling user $userName`: $_" -Level "ERROR"
      }
  }
  
  # Attach event handlers
  $refreshButton.Add_Click({ Refresh-Users })
  $enableButton.Add_Click({ Enable-SelectedUser })
  $disableButton.Add_Click({ Disable-SelectedUser })
  
  # Populate the grid with initial data
  Populate-Grid -data $users
  
  Write-Host "`nUser management window opened. You can now:" -ForegroundColor Cyan
  Write-Host "  - View user details in the grid" -ForegroundColor White
  Write-Host "  - Select a user and click 'Enable User' or 'Disable User'" -ForegroundColor White
  Write-Host "  - Click 'Refresh Users' to update the list" -ForegroundColor White
  Write-Host "Logs are being saved to: $env:TEMP\ADUserManager.log" -ForegroundColor White
  Write-Host "`nUse the grid to sort by any column by clicking the column header." -ForegroundColor Cyan
  
  # Display the form
  $form.ShowDialog()
  
  Write-Log "Application closed"
}

