Param (
    $PIDLStringDataFile = "$PSScriptRoot\PIDLStringCollection.xml"
)

####### Functions for managing various PIDL strings
# Since I don't currently have a way of Encoding PIDL strings, I have to manually gather the values from 
# either registry entries or INI files wherever they are stored.
# The PIDLs for ScreenSaver differ from Desktop and LockScreen folder. I guess they are a slightly different
# format.

$KnownPIDLStrings = $null

function Get-PIDLStringForPath {
    Param([string]$path)
    if ($KnownPIDLStrings -eq $null) {
        ImportDataCollection
    }    
    $KnownPIDLStrings.$path
}

####### ScreenSaver Slideshow Section Section #######
function Get-ScreenSaverSlideshowPIDLString  {

    Write-TraceCall -msg "$($MyInvocation.InvocationName)"

    $ScreenSaverRegSetting = @{
        Path="HKCU:\SOFTWARE\Microsoft\Windows Photo Viewer\Slideshow\Screensaver"
        Name='EncryptedPIDL'
    }
    (Get-ItemProperty -Path @ScreenSaverRegSetting).($ScreenSaverRegSetting.Name)

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}


####### Desktop Slideshow Section Section #######
function Get-DesktopSlideshowPIDLString {    
    Param ($iniFilePath = "$($env:userprofile)\AppData\Roaming\Microsoft\Windows\Themes\slideshow.ini")

    Write-TraceCall -msg "$($MyInvocation.InvocationName)"

    $Pairs = (Get-Content $iniFilePath | Select-String 'ImagesRootPIDL').ToString() | ConvertFrom-String -Delimiter '=' -PropertyNames 'Name','Value'
    ($Pairs | ? Name -like 'ImagesrootPIDL').Value

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

####### Lock Screen Slideshow Section #######
# Lock screen display depends on the screen saver timeout as well

function Get-LockSlideshowPIDLString {

    Write-TraceCall -msg "$($MyInvocation.InvocationName)"

    $LockScreenRegSetting = @{
        Path= "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Lock Screen"
        Name='SlideshowDirectoryPath1'
    }
    (Get-ItemProperty @LockScreenRegSetting).($LockScreenRegSetting.Name)

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}


function GatherData {
    Param (
        [Parameter (Mandatory=$true)]
        [ValidateSet('ScreenSaver','Desktop','LockScreen')]
        $source
    )

    $getters = @{
        'ScreenSaver'= get-item function:Get-ScreenSaverSlideshowPIDLString
        'Desktop'=     get-item function:Get-DesktopSlideshowPIDLString
        'LockScreen'=  get-item function:Get-LockSlideshowPIDLString
    }

    $plainPath = Get-Clipboard -Format Text -TextFormatType Text    
    $pidlPath = & $getters[$source]

    $KnownPIDLStrings.Add($plainPath, $pidlPath)
}

function ExportDataCollection {
    $KnownPIDLStrings | Export-Clixml -Path $PIDLStringDataFile
}

function ImportDataCollection {
    $KnownPIDLStrings = Import-Clixml -Path $PIDLStringDataFile
}
