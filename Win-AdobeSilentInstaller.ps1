$ProgressPreference="SilentlyContinue"

# The installer only runs when Adobe isn't running
# This kills all Adobe processes.
function _StopAdobeProcess{
    # Kill the Adobe process
    $ids=(Get-Process -Name acro*).id
    foreach ($id in $ids){
        Stop-Process -Force $id 2>&1 | Out-Null
        Wait-Process -Id $id 2>&1 | Out-Null
    }
}

# Fetch the uninstall string for adobe
function _UninstallAdobe{
    # Fetch the uninstall string for Adobe and grab the product id of it
    $progs=Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty
    # Uninstalling it silently
    try{
        Write-Host "Uninstalling Adobe if it's installed."
        (($progs | Where {$_.displayname -match "Adobe Acrobat.*DC"}).uninstallstring) -match "{.*.}" | Out-Null
        Start-Process "msiexec.exe" -ArgumentList "/X $($matches[0]) /qn" -Wait
    } catch {
        "Couldn't find Adobe Acrobat Reader. Installing it."
    }
}

# Installs Adobe using a hard coded link
# The MD5 sum is checked once. If it fails twice, Adobe is installed regardless
function _InstallAdobe{
    Write-Host "Downloading Adobe."
    Invoke-WebRequest -Uri 'https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2200120085/AcroRdrDC2200120085_en_US.exe' -OutFile "$env:temp\adobereader.exe" -UseBasicParsing
    $hashItShouldBe="3656695886E2E7A62A6C321DA1A22593"
    $hash=Get-FileHash $env:temp\adobereader.exe -Algorithm MD5
    if ($hash.hash -ne $hashItShouldBe){
      Write-Host "MD5Sum didn't match. Redownloading Adobe."
      Write-Host "The installer will run regardless if the MD5Sum doesn't match a second time."
      Invoke-WebRequest -Uri 'https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2200120085/AcroRdrDC2200120085_en_US.exe' -OutFile "$env:temp\adobereader.exe" -UseBasicParsing
    }
    Write-Host "Installing Adobe."
    Start-Process "$env:temp\adobereader.exe" -ArgumentList "/sAll /rs /msi EULA_ACCEPT=YES" -Wait
}

# Scrapes the DC release page to download the latest patch file
function _InstallLatestPatch{
    Write-Host "Locating the latest Adobe patch file for this month."
    $base="https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/"
    $link=(Invoke-WebRequest -Uri "https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/index.html" -UseBasicParsing).links.href | Select-String "continuous/dccontinuous" | Select -First 1
    $link = $base+$link

    Write-Host "Downloading the latest Adobe patch file."
    # After finding the patch notes for the month, let's parse out the URL for the latest patch file
    $patchlink=(Invoke-WebRequest -Uri "$link" -UseBasicParsing).links.href | Select-String "pub/adobe/reader/win/AcrobatDC/" | Select -First 1
    # Get the filename from the url and download it
    $filename=([uri]"$patchlink").Segments[-1]
    # Download the patch.msp
    Invoke-WebRequest -Uri "$patchlink" -OutFile "$env:temp/$filename" -UseBasicParsing
    Write-Host "Installing the patch."
    Start-Process "$env:temp/$filename" -ArgumentList "/qn" -Wait
}

# Scrapes the Adobe page to download the latest Extended Language Font Pack
function _InstallExtendedFontPack{
    $ProgressPreference="SilentlyContinue"
    $url="https://helpx.adobe.com/in/acrobat/kb/windows-font-packs-32-bit-reader.html"
    $link=(Invoke-WebRequest -Uri $url -UseBasicParsing).links.href | Select-String "pub/adobe/reader/win/AcrobatDC/misc/FontPack" | Select-Object -First 1
    Invoke-WebRequest -Uri "$link" -OutFile "$env:temp\ExtendedFontPack.msi" -UseBasicParsing

    Write-Host "Installing the Extended Font Pack"
    Start-Process "$env:temp\ExtendedFontPack.msi" -ArgumentList "/qn" -Wait
    Write-Host "DONE!"
}

# Disables automatic update through a mix of registry keys, scheduled tasks, and services
function _AdobeSettings{
    Write-Host "Disabling the update task, update services, update registry keys, and enhanced security mode."
    # Prevent autoupdates (not necessary with the enterprise edition)
    Disable-ScheduledTask -TaskName "Adobe Acrobat Update Task"
    Stop-Service -Name "Adobe Acrobat Update Service"
    Set-Service -Name "AdobeARMservice" -StartupType Disabled
    # Disables and locks down the update button
    New-Item "HKLM:\Software\WOW6432Node\Policies\Adobe\Acrobat Reader\DC" -ErrorAction SilentlyContinue
    New-Item "HKLM:\Software\WOW6432Node\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -ErrorAction SilentlyContinue
    New-ItemProperty "HKLM:\Software\WOW6432Node\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "Mode" -Value 0 -PropertyType DWORD -Force
    New-ItemProperty "HKLM:\Software\WOW6432Node\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "bUpdater" -Value 0 -PropertyType DWORD -Force
    New-ItemProperty "HKLM:\Software\WOW6432Node\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "bProtectedModeValue" -Value 0 -PropertyType DWORD -Force
    # This *should* disable enhanced security mode
    New-ItemProperty "HKLM:\Software\WOW6432Node\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "bEnhancedSecurityStandalone" -Value 0 -PropertyType DWORD -Force
    New-ItemProperty "HKLM:\Software\WOW6432Node\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "bEnhancedSecurityInBrowser" -Value 0 -PropertyType DWORD -Force
}

function _RemoveDesktopIcon{
    Remove-Item -Force -Path "C:\Users\*\Desktop\Acrobat Reader DC.lnk"
}

_UninstallAdobe
_InstallAdobe
_InstallLatestPatch
_InstallExtendedFontPack
_AdobeSettings
_RemoveDesktopIcon
