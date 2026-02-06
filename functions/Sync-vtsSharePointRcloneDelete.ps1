function Sync-vtsSharePointRcloneDelete {
    <#
    .SYNOPSIS
        Deletes files/folders from SharePoint destination that don't exist on source.
    
    .DESCRIPTION
        Compares source (Windows file share) with destination (SharePoint/OneDrive),
        identifies orphaned items on destination, and deletes them efficiently by
        purging top-level directories first.
    
    .PARAMETER Source
        Source path (Windows file share), e.g., "\\server2023\Network\public"
    
    .PARAMETER Destination
        Destination rclone remote, e.g., "PublicDocuments:"
    
    .PARAMETER DryRun
        Preview what would be deleted without actually deleting.
    
    .PARAMETER WorkingDir
        Directory for temporary files. Defaults to current directory.
    
    .PARAMETER UserAgent
        User-agent string for SharePoint. Defaults to recommended format.
    
    .EXAMPLE
        Sync-vtsSharePointRcloneDelete -Source "\\server2023\Network\public" -Destination "PublicDocuments:" -DryRun
    
    .EXAMPLE
        Sync-vtsSharePointRcloneDelete -Source "\\server2023\Network\public" -Destination "PublicDocuments:"

    .LINK
    Utilities
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Source,
        
        [Parameter(Mandatory=$true)]
        [string]$Destination,
        
        [switch]$DryRun,
        
        [string]$WorkingDir = (Get-Location),
        
        [string]$UserAgent = "NONISV|YourCompany|rclone/1.68.0"
    )
    
    # Create unique prefix for temp files
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $prefix = ($Destination -replace '[:\/]', '_') + "_$timestamp"
    
    $sourceFile = Join-Path $WorkingDir "${prefix}_source.txt"
    $destFile = Join-Path $WorkingDir "${prefix}_dest.txt"
    $toDeleteFile = Join-Path $WorkingDir "${prefix}_to_delete.txt"
    $topLevelDirsFile = Join-Path $WorkingDir "${prefix}_top_level_dirs.txt"
    $rootFilesFile = Join-Path $WorkingDir "${prefix}_root_files.txt"
    
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "SharePoint Cleanup: $Destination" -ForegroundColor Cyan
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Step 1: List source
    Write-Host "[1/6] Listing source: $Source" -ForegroundColor Yellow
    rclone lsf -R $Source > $sourceFile
    $sourceCount = (Get-Content $sourceFile).Count
    Write-Host "      Source items: $sourceCount" -ForegroundColor Green
    
    # Step 2: List destination
    Write-Host "[2/6] Listing destination: $Destination" -ForegroundColor Yellow
    rclone lsf -R $Destination --user-agent $UserAgent > $destFile
    $destCount = (Get-Content $destFile).Count
    Write-Host "      Destination items: $destCount" -ForegroundColor Green
    
    # Step 3: Find items to delete (on dest but not on source)
    Write-Host "[3/6] Comparing..." -ForegroundColor Yellow
    $sourceItems = Get-Content $sourceFile
    $destItems = Get-Content $destFile
    
    $toDelete = $destItems | Where-Object { $sourceItems -notcontains $_ }
    
    if ($toDelete.Count -eq 0) {
        Write-Host "      Nothing to delete! Destination is clean." -ForegroundColor Green
        Write-Host ""
        # Cleanup temp files
        Remove-Item $sourceFile, $destFile -ErrorAction SilentlyContinue
        return
    }
    
    $toDelete | Set-Content $toDeleteFile
    Write-Host "      Items to delete: $($toDelete.Count)" -ForegroundColor Magenta
    
    # Step 4: Identify top-level directories
    Write-Host "[4/6] Identifying top-level directories..." -ForegroundColor Yellow
    $dirs = $toDelete | Where-Object { $_ -match '/$' }
    $files = $toDelete | Where-Object { $_ -notmatch '/$' }
    
    # Find top-level dirs (no parent in delete list)
    $topLevelDirs = $dirs | Where-Object {
        $current = $_
        $parent = ($current -replace '[^/]+/$', '')
        ($parent -eq '') -or ($dirs -notcontains $parent)
    }
    
    # Find root-level files (no / in path)
    $rootFiles = $files | Where-Object { $_ -notmatch '/' }
    
    if ($topLevelDirs) { $topLevelDirs | Set-Content $topLevelDirsFile }
    if ($rootFiles) { $rootFiles | Set-Content $rootFilesFile }
    
    Write-Host "      Top-level directories: $($topLevelDirs.Count)" -ForegroundColor Green
    Write-Host "      Root-level files: $($rootFiles.Count)" -ForegroundColor Green
    
    # Step 5: Preview
    Write-Host "[5/6] Preview of items to delete:" -ForegroundColor Yellow
    Write-Host ""
    
    if ($topLevelDirs.Count -gt 0) {
        Write-Host "      Directories:" -ForegroundColor Cyan
        $topLevelDirs | Select-Object -First 10 | ForEach-Object {
            Write-Host "        - $_" -ForegroundColor Gray
        }
        if ($topLevelDirs.Count -gt 10) {
            Write-Host "        ... and $($topLevelDirs.Count - 10) more" -ForegroundColor Gray
        }
    }
    
    if ($rootFiles.Count -gt 0) {
        Write-Host "      Root files:" -ForegroundColor Cyan
        $rootFiles | Select-Object -First 10 | ForEach-Object {
            Write-Host "        - $_" -ForegroundColor Gray
        }
        if ($rootFiles.Count -gt 10) {
            Write-Host "        ... and $($rootFiles.Count - 10) more" -ForegroundColor Gray
        }
    }
    Write-Host ""
    
    # Step 6: Execute or dry-run
    if ($DryRun) {
        Write-Host "[6/6] DRY RUN - No changes made" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "To execute, run without -DryRun flag:" -ForegroundColor Cyan
        Write-Host "  Sync-vtsSharePointRcloneDelete -Source `"$Source`" -Destination `"$Destination`"" -ForegroundColor White
    } else {
        Write-Host "[6/6] Executing deletion..." -ForegroundColor Yellow
        
        # Purge directories
        if ($topLevelDirs.Count -gt 0) {
            Write-Host "      Purging $($topLevelDirs.Count) directories..." -ForegroundColor Cyan
            $dirCount = 0
            foreach ($dir in $topLevelDirs) {
                $dirCount++
                $dirPath = $dir.TrimEnd('/')
                Write-Progress -Activity "Purging directories" -Status "$dirPath" -PercentComplete (($dirCount / $topLevelDirs.Count) * 100)
                
                rclone purge "${Destination}${dirPath}" --user-agent $UserAgent --retries 3 --retries-sleep 10s 2>&1 | Out-Null
            }
            Write-Progress -Activity "Purging directories" -Completed
            Write-Host "      Directories purged." -ForegroundColor Green
        }
        
        # Delete root files
        if ($rootFiles.Count -gt 0) {
            Write-Host "      Deleting $($rootFiles.Count) root files..." -ForegroundColor Cyan
            rclone delete $Destination --files-from $rootFilesFile --user-agent $UserAgent --retries 3 --retries-sleep 10s
            Write-Host "      Root files deleted." -ForegroundColor Green
        }
        
        Write-Host ""
        Write-Host "=================================================" -ForegroundColor Green
        Write-Host "COMPLETE: Deleted $($toDelete.Count) items from $Destination" -ForegroundColor Green
        Write-Host "=================================================" -ForegroundColor Green
    }
    
    # Cleanup temp files (keep for dry-run review)
    if (-not $DryRun) {
        Remove-Item $sourceFile, $destFile, $toDeleteFile -ErrorAction SilentlyContinue
        Remove-Item $topLevelDirsFile, $rootFilesFile -ErrorAction SilentlyContinue
    } else {
        Write-Host ""
        Write-Host "Temp files saved for review:" -ForegroundColor Gray
        Write-Host "  $toDeleteFile" -ForegroundColor Gray
        Write-Host "  $topLevelDirsFile" -ForegroundColor Gray
        Write-Host "  $rootFilesFile" -ForegroundColor Gray
    }
}
