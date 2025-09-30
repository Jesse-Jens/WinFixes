#
# Script Name: Install-TeamsMeetingAddIn.ps1
#
# Synopsis:
#   Downloads the Microsoft Teams Meeting Add-in package from a specified URL,
#   extracts it, installs the add‑in folder in the proper location under
#   `%LOCALAPPDATA%\Microsoft`, and runs the MSI to repair the add‑in silently if
#   possible.  If the silent repair fails, it launches the installer
#   interactively.
#
# Description:
#   Many organizations distribute the Teams Meeting Add‑in for Outlook via an
#   MSI package packaged alongside a folder called `TeamsMeetingAdd‑in`.  In
#   some cases the add‑in becomes corrupted or missing from a user's machine.
#   This PowerShell script automates the remediation process by pulling a ZIP
#   archive from a GitHub repository (or any HTTP URL), extracting its
#   contents, copying the `TeamsMeetingAdd‑in` folder to the user's
#   `%LOCALAPPDATA%\Microsoft` directory, and then executing the MSI to
#   perform a full repair.  A silent repair is attempted first; if the
#   installer returns a non‑zero exit code, the script falls back to
#   launching the MSI without quiet flags so the user can complete the
#   repair manually.
#
# This version is designed for use with a single-command PowerShell invocation (e.g. using
# `irm`). It assumes that the `TeamsMeetingAddin.zip` archive resides in this repository
# and therefore hard‑codes the download URL so no parameters are required when running
# the script.
#
# To run this script via a single command, use:
#     irm https://raw.githubusercontent.com/Jesse-Jens/WinFixes/main/install-teamsmeeting-addin.ps1 | iex
#
# The script will download `TeamsMeetingAddin.zip` from this repository, extract it,
# copy the `TeamsMeetingAdd‑in` folder into `%LOCALAPPDATA%\Microsoft`, and repair the
# MSI silently if possible. See the README.md in the repository for more details.
#
# Notes:
#       * Requires PowerShell 5.1 or later.
#       * Run from an elevated prompt when possible.  Without administrative
#         privileges the MSI may not be able to repair the add‑in for all
#         users.

# Download URL for the Teams Meeting Add-in archive. Adjust this value if the zip file is stored elsewhere.
$ZipUrl = 'https://raw.githubusercontent.com/Jesse-Jens/WinFixes/main/TeamsMeetingAddin.zip'
# Generate a unique temporary working directory
$guid       = [System.Guid]::NewGuid().ToString()
$tempRoot   = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "TeamsAddIn_$guid")
$zipFile    = [System.IO.Path]::Combine($tempRoot, "TeamsMeetingAddin.zip")
$extractDir = [System.IO.Path]::Combine($tempRoot, "Extracted")

try {
    # Ensure working directories exist
    New-Item -Path $tempRoot   -ItemType Directory -Force | Out-Null
    New-Item -Path $extractDir -ItemType Directory -Force | Out-Null

    Write-Host "Downloading Teams Meeting Add-in package from $ZipUrl…" -ForegroundColor Cyan
    Invoke-WebRequest -Uri $ZipUrl -OutFile $zipFile

    Write-Host "Extracting package…" -ForegroundColor Cyan
    Expand-Archive -Path $zipFile -DestinationPath $extractDir -Force

    # Attempt to locate the 'TeamsMeetingAdd-in' folder in the extracted content.
    $sourceFolder = Get-ChildItem -Path $extractDir -Directory | Where-Object {
        $_.Name -match '^TeamsMeetingAdd\-?in'
    } | Select-Object -First 1
    if (-not $sourceFolder) {
        throw "Could not locate a folder named 'TeamsMeetingAdd-in' in the extracted archive."
    }

    # Prepare destination under %LOCALAPPDATA%\Microsoft
    $destRoot   = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Microsoft"
    if (-not (Test-Path -Path $destRoot)) {
        New-Item -Path $destRoot -ItemType Directory -Force | Out-Null
    }
    $destFolder = Join-Path -Path $destRoot -ChildPath $sourceFolder.Name

    # Remove any existing add-in folder before copying
    if (Test-Path -Path $destFolder) {
        Write-Host "Existing add-in folder detected at $destFolder; removing it…" -ForegroundColor Yellow
        Remove-Item -Path $destFolder -Recurse -Force
    }

    Write-Host "Copying add-in files to $destFolder…" -ForegroundColor Cyan
    Copy-Item -Path $sourceFolder.FullName -Destination $destFolder -Recurse -Force

    # Find the first MSI file in the extracted directory
    $msi = Get-ChildItem -Path $extractDir -Filter "*.msi" -File | Select-Object -First 1
    if ($msi) {
        $msiPath = $msi.FullName
        Write-Host "Attempting silent repair of the add-in using $msiPath…" -ForegroundColor Cyan

        $msiArgs  = "/fa \"$msiPath\" /quiet /norestart"
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -Wait -PassThru -ErrorAction SilentlyContinue

        # According to Windows Installer documentation, an exit code of 0 indicates success.
        if ($process.ExitCode -ne 0) {
            Write-Warning "Silent repair returned exit code $($process.ExitCode). Launching installer interactively…"
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i \"$msiPath\"" -Wait
        } else {
            Write-Host "Silent repair completed successfully." -ForegroundColor Green
        }
    } else {
        Write-Warning "No MSI installer found in the extracted package. Only files were copied."
    }

} catch {
    Write-Error "An error occurred during installation: $($_.Exception.Message)"
} finally {
    # Clean up temporary files and directories
    if (Test-Path -Path $tempRoot) {
        try { Remove-Item -Path $tempRoot -Recurse -Force } catch { }
    }
}