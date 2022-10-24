## Parameter set in Ninja-RMM

## Set Version Parameters
param ([string]$RequiredVersion = '')

## For Testing only
#$RequiredVersion = '22.003.20263'

## Check to make sure parameter is set
if ($RequiredVersion -eq '')
{ 
    Write-Host "Parameter is null. Please add the correct verson to the Parameter field."
    break 
}

## RegEx to remove periods from version. 

$RequiredVersionStrip = $RequiredVersion -replace '[.]'

## Check Access to FileServer

$path = "\\punto\Apps\Adobe\"
$accesscondition = Test-Path -Path $path
if ( $accesscondition )
{
    Write-Output "Computer has access to the fileshare."
}
    else {
        Write-Host "Check network access to the fileshare or try again."
        exit
    }

## Check Access to SCUT

$scutpath = "C:\scut\Adobe"
$accesscondition = Test-Path -Path $scutpath
if ( $accesscondition )
{
    Write-Output "Deployment directory already exists"
    cd $scutpath
}
    else {
        Write-Host "Creating SCUT deployment directory"
        mkdir C:\scut\adobe
        cd $scutpath
    }


## Check for existing msp files before downloading

$mspfiletest = "\\punto\Apps\Adobe\"+"*"+"$RequiredVersionStrip"+".msp"
$accesscondition = Test-Path -Path $mspfiletest
if ( $accesscondition )
{
    Write-Output "Adobe .MSP files exist. Skipping download."
}
    else {
    write-host "Correct version .MSP files not found. Please download correct version and specify in the attribute."
    break
        #$req = Invoke-WebRequest -uri "https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/continuous/dccontinuousoct2022.html#dccontinuousocttwentytwentytwo"
        #$req.Links | Where-Object {$_.href -like "*.msp"} | Select -ExpandProperty href
        #Start-BitsTransfer -Source $req -Destination \\punto\Apps\Adobe\
       }

## Get installed Adobe Products

$AdobeAcrobat64bit = Get-WmiObject -Class Win32_Product | Where-Object Caption -eq 'Adobe Acrobat (64-bit)' | Select-Object -ExpandProperty Caption

$AdobeAcrobatDC64bit = Get-WmiObject -Class Win32_Product | Where-Object Caption -eq 'Adobe Acrobat DC 64-bit' | Select-Object -ExpandProperty Caption

$AdobeAcrobatReader = Get-WmiObject -Class Win32_Product | Where-Object Caption -eq 'Adobe Acrobat Reader' | Select-Object -ExpandProperty Caption

$AdobeAcrobatReaderDC = Get-WmiObject -Class Win32_Product | Where-Object Caption -eq 'Adobe Acrobat Reader DC' | Select-Object -ExpandProperty Caption

$AdobeAcrobat = Get-WmiObject -Class Win32_Product | Where-Object Caption -eq 'Adobe Acrobat' | Select-Object -ExpandProperty Caption

## Check Adobe Product Versions 

$AdobeAcrobat64bitversion = Get-WmiObject -Class Win32_Product | Where-Object Name -eq 'Adobe Acrobat (64-bit)' | Select-Object -ExpandProperty Version
$CheckAdobeAcrobat64 = $AdobeAcrobat64bitversion -ge $RequiredVersion

$AdobeAcrobatDC64bitVersion = Get-WmiObject -Class Win32_Product | Where-Object Name -eq 'Adobe Acrobat DC 64-bit)' | Select-Object Version | Select-Object -ExpandProperty Version
$CheckVersionAcrobatDC64 = $AdobeAcrobatDC64bitVersion -ge $RequiredVersion

$AdobeAcrobatReaderVersion = Get-WmiObject -Class Win32_Product | Where-Object Name -eq 'Adobe Acrobat Reader' | Select-Object Version | Select-Object -ExpandProperty Version
$CheckVersionAcrobatReader = $AdobeAcrobatReaderVersion -ge $RequiredVersion

$AdobeAcrobatReaderDCVersion = Get-WmiObject -Class Win32_Product | Where-Object Name -eq 'Adobe Acrobat Reader DC' | Select-Object Version | Select-Object -ExpandProperty Version
$CheckVersionAcrobatReader32DC = $AdobeAcrobatReaderDCVersion -ge $RequiredVersion

$AbobeAcrobatversion = Get-WmiObject -Class Win32_Product | Where-Object Name -eq 'Adobe Acrobat' | Select-Object Version | Select-Object -ExpandProperty Version
$CheckVersionAcrobat = $AbobeAcrobatversion -ge $RequiredVersion

## Adobe Acrobat 64 bit Update

if ($null -eq $AdobeAcrobat64bit ) {
    Write-Host "Adobe Acrobat 64-bit is not installed"
    }
        elseif ($AdobeAcrobat64bit -eq $AdobeAcrobat64bit -and $CheckAdobeAcrobat64 -eq "true") {
                Write-host "Adobe Acrobat 64 bit installed. Adobe is up to date."
        }
            elseif ($AdobeAcrobat64bit -eq $AdobeAcrobat64bit -and $CheckAdobeAcrobat64 -ne "true") {
                    write-host "Adobe Acrobat 64 bit Installed, but needs update. Starting Update"
                    cp $path"AcrobatDCx64Upd"$requiredversionstrip".msp"
                    msiexec.exe /p "AcrobatDCx64Upd$requiredversionstrip.msp" /qn
              }


## Abobe Acrobat DC 64-bit

if ($null -eq $AdobeAcrobatDC64bit ) {
    Write-Host "Adobe Acrobat DC 64-bit is not installed"
    }
        elseif ($AdobeAcrobatDC64bit -eq $AdobeAcrobatDC64bit -and $CheckVersionAcrobatDC64 -eq "true") {
                Write-host "Adobe Acrobat DC 64 bit installed. Adobe is up to date."
        }
            elseif ($AdobeAcrobatDC64bit -eq $AdobeAcrobatDC64bit -and $CheckVersionAcrobatDC64 -ne "true") {
                    write-host "Adobe Acrobat DC 64 bit Installed, but needs update. Starting Update"
                    cp $path"AcrobatDCx64Upd"$requiredversionstrip".msp"
                    msiexec.exe /p "AcrobatDCx64Upd$requiredversionstrip.msp" /qn               
              }

## Adobe Acrobat Reader

if ($null -eq $AdobeAcrobatReader ) {
    Write-Host "Adobe Acrobat Reader is not installed"
    }
        elseif ($AdobeAcrobatReader -eq $AdobeAcrobatReader -and $CheckVersionAcrobatReader -eq "true") {
                Write-host "Adobe Acrobat Reader installed. Adobe is up to date."
        }
            elseif ($AdobeAcrobatReader -eq $AdobeAcrobatReader -and $CheckVersionAcrobatReader -ne "true") {
                    write-host "Adobe Acrobat Reader Installed, but needs update. Starting Update"
                    cp $path"AcroRdrDCUpd"$requiredversionstrip".msp"
                    msiexec.exe /p "AcroRdrDCUpd$requiredversionstrip.msp" /qn
              }

## Adobe Acrobat Reader DC

if ($null -eq $AdobeAcrobatReaderDC ) {
    Write-Host "Adobe Acrobat Reader DC is not installed"
    }
        elseif ($AdobeAcrobatReaderDC -eq $AdobeAcrobatReaderDC -and $CheckVersionAcrobatReader32DC -eq "true") {
                Write-host "Adobe Acrobat Reader DC installed. Adobe is up to date."
        }
            elseif ($AdobeAcrobatReaderDC -eq $AdobeAcrobatReaderDC -and $CheckVersionAcrobatReader32DC -ne "true") {
                    write-host "Adobe Acrobat Reader DC Installed, but needs update. Starting Update"
                    cp $path"AcroRdrDCUpd"$requiredversionstrip".msp"
                    msiexec.exe /p "AcroRdrDCUpd$requiredversionstrip.msp" /qn
              }

## Adobe Acrobat

if ($null -eq $AdobeAcrobat ) {
    Write-Host "Adobe Acrobat is not installed"
    }
        elseif ($AdobeAcrobat -eq $AdobeAcrobat -and $CheckVersionAcrobat -eq "true") {
                Write-host "Adobe Acrobat installed. Adobe is up to date."
        }
            elseif ($AdobeAcrobat -eq $AdobeAcrobat -and $CheckVersionAcrobat -ne "true") {
                    write-host "Adobe Acrobat Installed, but needs update. Starting Update"
                    cp $path"AcrobatDCUpd"$requiredversionstrip".msp"
                    msiexec.exe /p "AcrobatDCUpd$requiredversionstrip.msp" /qn
              }

## Standby for install to complete
 
Start-Sleep -Seconds 300

## Cleanup scut Folder

cd C:\
Remove-Item $scutpath -Force -Recurse

## Display Current Versions of Adobe

Write-Host "Current Adobe Acrobat (64-bit) Version"
Get-WmiObject -Class Win32_Product | Where-Object Name -eq 'Adobe Acrobat (64-bit)' | Select-Object -ExpandProperty Version

Write-Host "Current Adobe Acrobat DC 64-bit Version"
Get-WmiObject -Class Win32_Product | Where-Object Name -eq 'Adobe Acrobat DC 64-bit)' | Select-Object Version | Select-Object -ExpandProperty Version

Write-Host "Current Adobe Acrobat Reader Version"
Get-WmiObject -Class Win32_Product | Where-Object Name -eq 'Adobe Acrobat Reader' | Select-Object Version | Select-Object -ExpandProperty Version

Write-Host "Current Adobe Acrobat Reader DC Version"
Get-WmiObject -Class Win32_Product | Where-Object Name -eq 'Adobe Acrobat Reader DC' | Select-Object Version | Select-Object -ExpandProperty Version

Write-Host "Current Adobe Acrobat Version"
Get-WmiObject -Class Win32_Product | Where-Object Name -eq 'Adobe Acrobat' | Select-Object Version | Select-Object -ExpandProperty Version

Write-Host "Update Complete"
