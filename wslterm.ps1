# Inspiration from 
# Silent Install 7-Zip
# http://www.7-zip.org/download.html
# https://forum.pulseway.com/topic/1939-install-7-zip-with-powershell/ 

# Check for admin rights
$wid = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$prp = new-object System.Security.Principal.WindowsPrincipal($wid)
$adm = [System.Security.Principal.WindowsBuiltInRole]::Administrator
if (-not $prp.IsInRole($adm)) {
    throw "This script requires elevated rights to install software.. Please run from an elevated shell session."
}

# Check for 7z install
Write-Progress -Activity "Validating Dependencies" -Status "Checking for 7zip"
$7z_Application = get-command 7z.exe -ErrorAction SilentlyContinue | select-object -expandproperty Path
if ([string]::IsNullOrEmpty($7z_Application)) {   
    $7z_Application = "C:\Program Files\7-Zip\7z.exe"
}

if (-not (Test-Path $7z_Application)) {
    Write-Progress -Activity "Validating Dependencies" -Status "Installing 7zip"
    # Path for the workdir
    $workdir = "c:\installer\"

    # Check if work directory exists if not create it
    If (-not (Test-Path -Path $workdir -PathType Container)) { 
        New-Item -Path $workdir  -ItemType directory 
    }

    # Download the installer
    $source = "http://www.7-zip.org/a/7z1801-x64.msi"
    $destination = "$workdir\7-Zip.msi"

    Invoke-WebRequest $source -OutFile $destination 

    # Start the installation
    msiexec.exe /i "$workdir\7-Zip.msi" /qb
    # Wait XX Seconds for the installation to finish
    Start-Sleep -s 35

    # Remove the installer
    Remove-Item -Force $workdir\7*
    Write-Progress -Activity "Validating Dependencies" -Status "Installing 7zip" -Completed	
}
Write-Progress -Activity "Validating Dependencies" -Completed

Write-Progress -Activity "Ensure in `$HOME directory"
set-location $env:USERPROFILE

# Set variable for WSL terminal
$version = "0.8.8"
$wslTerminal = "wsl-terminal-$version.7z"

Write-Progress -Activity "Get bits for WSL terminal"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest -Uri "https://github.com/goreliu/wsl-terminal/releases/download/v$version/$wslTerminal" -OutFile $env:USERPROFILE\$wslTerminal

Write-Progress -Activity "Extract WSL terminal and remove after complete"
Get-Item $wslTerminal | ForEach-Object {
    $7z_Arguments = @(
        'x'							## eXtract files with full paths
        '-y'						## assume Yes on all queries
        "`"-o$($env:USERPROFILE)`""		## set Output directory
        "`"$($_.FullName)`""				## <archive_name>
    )
    & $7z_Application $7z_Arguments
    If ($LASTEXITCODE -eq 0) {
        Remove-Item -Path $_.FullName -Force
    }
}

Write-Progress -Activity "Ensure symlink exists"
$symlink = "$env:USERPROFILE\Desktop\wsl.lnk"
If (-not (Test-Path -Path $symlink)) {
    New-Item -ItemType SymbolicLink -Path "$env:USERPROFILE\Desktop\" -Name "wsl.lnk" -Value "$env:USERPROFILE\wsl-terminal\open-wsl.exe" 
}
# Download the font
Install-Font

Function Install-Font {
    <# 
        .Synopsis 
        Installs one or more fonts. 
        .Parameter FontPath 
        The path to the font to be installed or a directory containing fonts to install. 
        .Parameter Recurse 
        Searches for fonts to install recursively when a path to a directory is provided. 
        .Notes 
        There's no checking if a given font is already installed. This is problematic as an existing 
        installation will trigger a GUI dialogue requesting confirmation to overwrite the installed 
        font, breaking unattended and CLI-only scenarios. 
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$FontPath,

        [Switch]$Recurse
    )

    $ErrorActionPreference = 'Stop'
    $ShellAppFontNamespace = 0x14

    if (Test-Path -Path $FontPath) {
        $FontItem = Get-Item -Path $FontPath
        if ($FontItem -is [IO.DirectoryInfo]) {
            if ($Recurse) {
                $Fonts = Get-ChildItem -Path $FontItem -Include ('*.fon','*.otf','*.ttc','*.ttf') -Recurse
            } else {
                $Fonts = Get-ChildItem -Path "$FontItem\*" -Include ('*.fon','*.otf','*.ttc','*.ttf')
            }

            if (!$Fonts) {
                throw ('Unable to locate any fonts in provided directory: {0}' -f $FontItem.FullName)
            }
        } elseif ($FontItem -is [IO.FileInfo]) {
            if ($FontItem.Extension -notin ('.fon','.otf','.ttc','.ttf')) {
                throw ('Provided file does not appear to be a valid font: {0}' -f $FontItem.FullName)
            }

            $Fonts = $FontItem
        } else {
            throw ('Expected directory or file but received: {0}' -f $FontItem.GetType().Name)
        }
    } else {
        throw ('Provided font path does not appear to be valid: {0}' -f $FontPath)
    }

    $ShellApp = New-Object -ComObject Shell.Application
    $FontsFolder = $ShellApp.NameSpace($ShellAppFontNamespace)
    foreach ($Font in $Fonts) {
        If (Test-Path "c:\windows\fonts\$($Font.FullName)") {
          throw ('Provided font: {0}' -f $FontItem.FullName exist in path: {0}' -f $FontPath)
        } else {
        Write-Verbose -Message ('Installing font: {0}' -f $Font.BaseName)
        $FontsFolder.CopyHere($Font.FullName)
        }
    }
}
