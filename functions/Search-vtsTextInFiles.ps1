function Search-vtsTextInFiles {
<#
.SYNOPSIS
Search recursively through a directory for text matches in probable text files.

.DESCRIPTION
Search-TextInFiles scans all files under the specified path, evaluating whether 
each file is likely to be a text file before reading it. It performs a 
line-by-line search for the specified string using either case-sensitive or 
case-insensitive comparison. For each file that contains matches, it outputs 
the file path and the line numbers where matches were found.

.PARAMETER Path
The root directory to begin searching from.

.PARAMETER SearchTerm
The text string to search for within each file.

.PARAMETER CaseSensitive
Switch to enable case-sensitive matching. By default the search is case-insensitive.

.EXAMPLE
Search-TextInFiles -Path "C:\Logs" -SearchTerm "error"
Searches all text files under C:\Logs (case-insensitive) and prints matches.

.EXAMPLE
Search-TextInFiles -Path . -SearchTerm "TokenExpired" -CaseSensitive
Performs case-sensitive search for "TokenExpired" in the current directory.

.NOTES
This function uses .NET APIs for efficient directory traversal and file access.
It can read files that are locked for writing by other processes.
#>

    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,

        [switch]$CaseSensitive
    )

    # Validate path
    if (-not (Test-Path -Path $Path)) {
        Write-Error "Path not found: $Path"
        exit 2
    }

    $comparison = if ($CaseSensitive.IsPresent) {
        [System.StringComparison]::Ordinal
    } else {
        [System.StringComparison]::OrdinalIgnoreCase
    }

    function Is-ProbablyTextFile {
        param(
            [string]$FilePath,
            [int]$SampleBytes = 8192
        )

        try {
            $fs = [System.IO.File]::Open(
                $FilePath,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::ReadWrite
            )
        } catch {
            return $false
        }

        try {
            $buffer = New-Object byte[] ([Math]::Min($SampleBytes, [int]$fs.Length))
            if ($buffer.Length -le 0) { $fs.Close(); return $false }

            $bytesRead = $fs.Read($buffer, 0, $buffer.Length)

            for ($i = 0; $i -lt $bytesRead; $i++) {
                if ($buffer[$i] -eq 0) {
                    $fs.Close()
                    return $false
                }
            }

            $controlCount = 0
            for ($i = 0; $i -lt $bytesRead; $i++) {
                $b = $buffer[$i]
                if ($b -lt 32 -and $b -ne 9 -and $b -ne 10 -and $b -ne 13) { 
                    $controlCount++ 
                }
            }

            $fs.Close()
            if ($controlCount -gt ($bytesRead * 0.3)) { return $false }

            return $true
        } catch {
            try { $fs.Close() } catch {}
            return $false
        }
    }

    # Directory traversal using .NET
    $dirStack = New-Object System.Collections.Generic.Stack[string]
    $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    $dirStack.Push($resolved)

    while ($dirStack.Count -gt 0) {
        $currentDir = $dirStack.Pop()

        try {
            $files = [System.IO.Directory]::GetFiles($currentDir)
        } catch {
            continue
        }

        foreach ($file in $files) {
            if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { continue }

            if (-not (Is-ProbablyTextFile -FilePath $file)) { continue }

            try {
                $fs = [System.IO.File]::Open(
                    $file,
                    [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read,
                    [System.IO.FileShare]::ReadWrite
                )
                $sr = New-Object System.IO.StreamReader($fs, $true)
            } catch {
                try { if ($fs) { $fs.Close() } } catch {}
                continue
            }

            try {
                $lineNumber = 0
                $matches = New-Object System.Collections.Generic.List[int]

                while (($line = $sr.ReadLine()) -ne $null) {
                    $lineNumber++
                    if ($line.IndexOf($SearchTerm, $comparison) -ge 0) {
                        $matches.Add($lineNumber)
                    }
                }

                if ($matches.Count -gt 0) {
                    $lines = ($matches | ForEach-Object { $_ }) -join ','
                    [Console]::WriteLine("{0}`tLines: {1}", $file, $lines)
                }
            } catch {
            } finally {
                try { $sr.Close() } catch {}
                try { $fs.Close() } catch {}
            }
        }

        try {
            $subdirs = [System.IO.Directory]::GetDirectories($currentDir)
            foreach ($d in $subdirs) {
                $dirStack.Push($d)
            }
        } catch {}
    }
}
