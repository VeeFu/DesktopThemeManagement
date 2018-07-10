Param (    
    $LogLevel = @('Debug','Trace'),
    $LogFile = "\\jakartafs01.eap.state.sbu\workstationInfo$\Logs\$env:computername.trace.log"
)

. $PSScriptRoot\Logging.ps1 -LogLevel $LogLevel -LogFile $LogFile

####### Functions for managing various PIDL strings
# Since I don't currently have a way of Encoding PIDL strings, I have to manually gather the values from 
# either registry entries or INI files wherever they are stored.
# The PIDLs for ScreenSaver differ from Desktop and LockScreen folder. I guess they are a slightly different
# format.
#
# copy target path to clipboard.
# make changes to the reference computer's settings
# run GatherData

function Get-PIDLStringForPath {
    Param([string]$path)
    Begin {

        Write-Trace -msg "Calling $($MyInvocation.InvocationName)" 

        $KnownPIDLStrings = Import-Clixml -Path $PSScriptRoot\PIDLStringCollection.xml
        #$KnownPIDLStrings = Import-Clixml -Path $PSScriptRoot\DataCollection.xml
        #$KnownPIDLStrings = Import-Clixml -Path \\jakartafp06\isc\Adminshell\drakevg\ScreenSaverSettings\dataCollection.xml
    }
    End {
        $KnownPIDLStrings.$path

        Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
    }
}

function GatherData {
    Param (
        [Parameter (Mandatory=$true)]
        [ValidateSet('ScreenSaver','Desktop','LockScreen')]
        $source
    )

    Write-TraceCall -msg "$($MyInvocation.InvocationName)"

    $getters = @{
        'ScreenSaver'= get-item function:Get-ScreenSaverSlideshowPIDLString
        'Desktop'=     get-item function:Get-DesktopSlideshowPIDLString
        'LockScreen'=  get-item function:Get-LockSlideshowPIDLString
    }

    $plainPath = Get-Clipboard -Format Text -TextFormatType Text    
    $pidlPath = & $getters[$source]

    $DataCollection.Add($plainPath, $pidlPath)

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

function ExportDataCollection {
    $DataCollection | Export-Clixml -Path "$srcdir\dataCollection.xml"
}

function ImportDataCollection {
    $DataCollection =  Import-Clixml -Path "$srcdir\dataCollection.xml"
}


####### Local state management for login script #######
$global:LocalStateData = $null
$global:LocalStateDataFile = "c:\ProgramData\ThemeManagement\stateData.xml"

function Load-LocalStateData {

    Write-TraceCall -msg "$($MyInvocation.InvocationName)"

    if ($global:LocalStateData -eq $null) {
        if (Test-Path $LocalStateDataFile) {
            $global:LocalStateData = Import-Clixml $LocalStateDataFile            
        }
        if ($global:LocalStateData -eq $null) {
            $global:LocalStateData = @{}
        }
    }
    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

function Save-LocalStateData {
    
    Write-Trace -msg "Calling $($MyInvocation.InvocationName)"
    
    Write-Host "data file path is: $global:LocalStateDataFile"
    $savetoDirectory = [System.IO.Path]::GetDirectoryName($global:LocalStateDataFile)
    if ((Test-Path $savetoDirectory) -eq $false ) {
        New-Item $savetoDirectory -Force -ItemType Container | Out-Null
    }
    $global:LocalStateData | Export-Clixml -Path $global:LocalStateDataFile

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

function Set-LastWritetimeRecord {
    Param(
        [String]$filepath,
        [DateTime]$lastwritetime
    )

    Write-TraceCall -msg "$($MyInvocation.InvocationName)"

    $cacheName = 'LastWriteTime'
    Load-LocalStateData
    if ($global:LocalStateData[$cacheName] -eq $null) {
        $global:LocalStateData[$cacheName] = @{}
    }
    $global:LocalStateData[$cacheName][$filepath] = $lastwritetime

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

function Get-LastWritetimeRecord {    
    Param ($filepath)

    Write-TraceCall -msg "$($MyInvocation.InvocationName)"

    $cacheName = 'LastWriteTime'
    Load-LocalStateData
    if ($global:LocalStateData[$cacheName] -eq $null) {
        $global:LocalStateData[$cacheName] = @{}
    }
    $global:LocalStateData[$cacheName][$filepath]

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

function Test-FilesHaveChanged {
    Param ($pathFilter)
    
    "Calling $($MyInvocation.InvocationName)" | Out-File -Append -FilePath $logfile

    [bool]$filesHaveChanged = $false
    Get-ChildItem -Path $pathFilter | ForEach-Object {
        $cachedLastWriteTime = Get-LastWritetimeRecord -filepath $_.FullName
        if ($cachedLastWriteTime -eq $null -or $_.LastWriteTime -ne $cachedLastWriteTime ){
            Set-LastWritetimeRecord -filepath $_.FullName -lastwritetime $_.LastWriteTime
            $filesHaveChanged = $true
        }
    }
    $filesHaveChanged

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
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

function Choose-BestDesktopSlideDimensions {
    Param ($RootPath)
    
    Write-TraceCall -msg "$($MyInvocation.InvocationName)"

    $candidates = (Get-ChildItem $RootPath -Directory).Name | 
        Select-Object @{label='Width';expression={$_ -replace "(\d+)x\d+",'$1'}},
                      @{label='Height';expression={$_ -replace "\d+x(\d+)",'$1'}}
    $bestFit = Get-BestFitDimensions -source (Get-PrimaryMonitorDimensions) -candidates $candidates
    "$RootPath\$($bestFit.Width)x$($bestfit.Height)"

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

function Set-DesktopSlideshow {
    Param (
        $iniFilePath = "$($env:userprofile)\AppData\Roaming\Microsoft\Windows\Themes\slideshow.ini",
        $PathToSlides
    )

    Write-TraceCall -msg "$($MyInvocation.InvocationName)"

    $currentSlideshowPathPIDL = Get-DesktopSlideshowPIDLString
    $newPIDLString = Get-PIDLStringForPath -path $PathToSlides

    # new desktop background slideshow folder is being set
    if ((Test-FilesHaveChanged -pathFilter $PathToSlides) -or $currentSlideshowPathPIDL -notlike $newPIDLString) {
        $inifile = Get-Item $iniFilePath -Force
        
        $newINIContent =
@"
[Slideshow]
ImagesRootPIDL=$newPIDLString
"@
        if ($inifile -ne $null) {
            $oldAttriutes = $inifile.Attributes
            $inifile.Attributes = 'Archive'
        } else {
            $oldAttributes = @('Archive','Hidden')
        }
        
        $newINIContent | out-file $iniFilePath -Force
        (get-item $iniFilePath).Attributes = $oldAttriutes
        
        # using this to force update of explorer with the new slideshow path, 
        # but it might be possible to trigger in a lighter-weight way
        Stop-Process -Name explorer
        Start-Sleep -Seconds 1
        if (-not (Get-Process -name Explorer)){
            explorer.exe
        }
    }

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

function Build-RegistryPathIfMissing {
    Param ($regKeyPath)

    Write-TraceCall -msg "$($MyInvocation.InvocationName)"

    $splitRegKey = $regKeyPath -split '\\'
    $pathIterator = 0
    while ($pathIterator -lt $splitRegKey.count) {
        $testRegPath = ($splitRegKey[0..$pathIterator]) -join '\'
        if ((Test-Path $testRegPath) -eq $false) {
            #Write-host $testRegPath
            New-Item $testRegPath
        }
        $pathIterator++
    }

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

function Set-LockScreenSlideshow {
    param ($PathToSlides)
    
    Write-TraceCall -msg "$($MyInvocation.InvocationName)"

    $hkcuLockScreen = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Lock Screen"
    $hkcuContentDeliveryManager = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    $khcuDesktop = "HKCU:\Control Panel\Desktop"

    $newPIDLPath = Get-PIDLStringForPath $PathToSlides

    if ($newPIDLPath -ne $null) {
        $LockScreenRegistrySettings = @(
        @{  # Enables/Disables lock screen slideshow
            Path=$hkcuLockScreen
            Name='SlideshowEnabled'
            Value=1
            PropertyType="DWord"
        },
        @{  # Contains binary encoded string for the lock screen image collection path
            Path=$hkcuLockScreen
            Name='SlideshowDirectoryPath1'
            Value=$newPIDLPath
            PropertyType="String"
        },
        @{  # controls whether hints and tips are enabled on the lock screen (leave disabled)
            Path=$hkcuContentDeliveryManager
            Name='RotatingLockScreenOverlayEnabled'
            Value=0
            PropertyType="DWord"
        },
        @{  # Seems to be redundant, but unsure. Set to 1 anyway
            Path=$hkcuContentDeliveryManager
            Name='RotatingLockScreenEnabled'
            Value=1
            PropertyType="DWord"
        },
        @{  # Do not include camera roll. We do not wish to accidentally show private pictures
            Path=$hkcuLockScreen
            Name='SlideshowIncludeCameraRoll'
            Value=0
            PropertyType="DWord"
        },
        @{  # Choose only images that fit the current screen resolution
            Path=$hkcuLockScreen
            Name="SlideshowOptimizePhotoSelection"
            Value=0
            PropertyType="DWord"
        },
        @{  # When inactive, show lock screen instead of turning off the screen
            Path=$hkcuLockScreen
            Name='SlideshowAutoLock'
            Value=1
            PropertyType="DWord"
        },
        @{  # How long to show the slideshow before turning off the screen
            Path=$hkcuLockScreen
            Name='SlideshowDuration'
            Value=1800000
            PropertyType="DWord"
        },
        @{  # timeout in seconds before Lock Screen Slideshow begins
            Path=$khcuDesktop
            Name='ScreenSaveTimeOut'
            Value=600
            PropertyType="DWord"
        }
        )

        $LockScreenRegistrySettings | %{
            Build-RegistryPathIfMissing -regkeypath $_.path
            if ((Get-ItemProperty -path $_.Path -Name $_.Name) -ne $_.Value) {
                Set-ItemProperty -Path $_.path -Name $_.Name -Value $_.Value
            } else {
                New-ItemProperty @_
            }
        }
    } else {
        Write-Error "Could not find $PathToSlides in the PIDL cache"
    }
    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

####### Function for inspecting image dimensions #######

function Get-ImageFileDimensions {
    Param (
        [Parameter (mandatory=$true,ValueFromPipeline=$true)]
        [System.IO.FileSystemInfo]
        $pathToImage
    )
    Begin {
        Write-TraceCall "$($MyInvocation.InvocationName)"
    }
    Process {
        Add-Type -AssemblyName System.Drawing
        $img = [System.Drawing.Image]::FromFile((get-item $pathToImage).fullname)
        [PSCustomObject] @{Width=$img.Width;Height=$img.Height}
        $img.Dispose()
    }
    End {
        Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
    }
}

####### Function for retrieving primary monitor dimensions #######

function Get-PrimaryMonitorDimensions {

    Write-TraceCall "$($MyInvocation.InvocationName)"

    Add-Type -AssemblyName System.Windows.Forms
    $PrimaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    [PSCustomObject] @{Width=$PrimaryScreen.Bounds.Width;Height=$PrimaryScreen.bounds.Height}

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

function Get-PrimaryMonitorAspectRatio {

    Write-TraceCall "$($MyInvocation.InvocationName)"

    $dims = Get-PrimaryMonitorDimensions 
    [float] $dims.Width / $dims.Height

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

function Get-BestFitDimensions {
    Param(
        [Object]   $source,
        [Object[]] $candidates
    )

    Write-TraceCall "$($MyInvocation.InvocationName)"

    #if there's an exact match, return it
    $exactMatch = $candidates | Where-Object {
        $_.Width -eq $source.Width -and $_.Height -eq $source.Height
    }

    if ($exactMatch) {
        $exactMatch | Select-Object -First 1
    } else {
        $sourceAspectRatio = $source.Width / $source.Height
        $calculations = $candidates | Select-Object Width,
            Height,
            @{label='ratioDifference'; Expression={[float][math]::Abs(($_.Width / $_.Height) - $sourceAspectRatio)}},
            @{label='size'; Expression={$_.Width * $_.Height}}
        
        #get minimum ratio difference across candidates
        $BestDiff = $calculations.RatioDifference | Sort-Object | Select-Object -First 1

        #select candidate with best ratio difference and highest resolution
        $calculations | Where-Object ratioDifference -like $BestDiff |
            Sort-Object -Property size -Descending |
            Select-Object -Property Width,Height |
            Select-Object -First 1
    }
    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

function Test-Get-BestFitDimesions {
    $source1 = [pscustomobject]@{Width=100;Height=200}
    $test1 = @(
        [pscustomobject]@{Width=50;Height=100},
        [pscustomobject]@{Width=100;Height=200},
        [pscustomobject]@{Width=200;Height=400}
    )    
    $expectedResult1 = [pscustomobject]@{Width=100;Height=200}
    
    $result1 = Get-BestFitDimensions -source $source1 -candidates $test1
    if ($result1 -notlike $expectedResult1) {
        Write-Host 'Get-BestFitDimensions failed exact-match test.'
    }

    $source2 = [pscustomobject]@{Width=100;Height=200}
    $test2 = @(
        [pscustomobject]@{Width=200;Height=400},
        [pscustomobject]@{Width=400;Height=800}
    )
    
    $expectedResult2 = [pscustomobject]@{Width=400;Height=800}
    $result2 = Get-BestFitDimensions -source $source2 -candidates $test2
    if ($result2 -notlike $expectedResult2) {
        Write-Host "Get-BestFitDimensions highest size when no exact-match found. Actually returned $($result2.Width) x $($result2.height)"
    }
    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}


####### Utility functions for calling into WinRT #######

Add-Type -AssemblyName System.Runtime.WindowsRuntime

$asTaskGeneric = (
    [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object { 
            $_.Name -eq 'AsTask' -and 
            $_.GetParameters().Count -eq 1 -and 
            $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' 
        }
)[0]

Function Await($WinRtTask, $ResultType) {
    Write-TraceCall -msg "$($MyInvocation.InvocationName)"
    
    Write-TraceCall -msg "MakeGenericMethod"
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)

    Write-TraceCall -msg "Calling Invoke"
    $netTask = $asTask.Invoke($null, @($WinRtTask))

    Write-TraceCall -msg "Calling Wait"
    $netTask.Wait(-1) | Out-Null
    
    $result = $netTask.Result
    Write-Debug "Getting Result: $($result)"

    Write-Output $result
    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

####### Function calls into WinRT to set Lock Screen #######

function Set-LockScreenImage {
    Param ($pathToTheme)

    Write-TraceCall "$($MyInvocation.InvocationName)"

    Add-Type -AssemblyName System.Windows.Forms
    [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime] | Out-Null
    [Windows.System.UserProfile.LockScreen,Windows.System.UserProfile,ContentType=WindowsRuntime] | Out-Null

    # Get candidate images named LockScreen*.jpg
    $lockScreenFileFilter = "$pathToTheme\LockScreen*.jpg"
    
    # lets set the lock screen regardless for now
    Test-FilesHaveChanged -pathFilter $lockScreenFileFilter
        
    $availableDimensions = Get-ChildItem $lockScreenFileFilter |
        Get-ImageFileDimensions |
        Select-Object -Unique Width,Height

    $bestDimensions = Get-BestFitDimensions -source (Get-PrimaryMonitorDimensions) -candidates $availableDimensions

    $newLockScreenPath = Get-ChildItem $lockScreenFileFilter | 
        Where-Object { ($_ | Get-ImageFileDimensions ) -like $bestDimensions } |
        Select-Object -First 1

    if (
        (Test-Path $newLockScreenPath ) -and 
        (
            (([Windows.System.UserProfile.LockScreen]::OriginalImageFile).OriginalString -notlike $newLockScreenPath) -or
            (Test-FilesHaveChanged -pathFilter $newLockScreenPath)
        )
    ) {            
        $newLockScreen = Await -WinRtTask ([Windows.Storage.StorageFile]::GetFileFromPathAsync($newLockScreenPath)) -ResultType ([Windows.Storage.StorageFile])
        "  Calling SetImageFileAsync" | Out-File -FilePath $logfile -Append
        [Windows.System.UserProfile.LockScreen]::SetImageFileAsync($newLockScreen) | Out-Null
    }
    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

####### Tests the named theme directory for requirements #######
# Called by the login script before it enacts changes
# Checks to ensure that the required paths exist in the PIDL path
# Checks that required directories and filenames exist
function Test-Theme {
    Param (
        [String]$themeName,
        [String]$localThemePath = "C:\users\Public\Public Pictures\Themes",
        [String]$networkThemePath = "\\jakartafs01.eap.state.sbu\DesktopThemes$"
    )

    Write-TraceCall "$($MyInvocation.InvocationName)"

    $ThemeSubdirs = @(
        'LockScreenSlides',
        'DesktopSlides',
        'DesktopSlides\1280x1024',
        'DesktopSlides\1680x1050')
    $ThemeFiles = @('LockScreen*.jpg')
    $status = 'Success'

    $ThemeSubdirs | ForEach-Object {
        $testresult = Get-PIDLStringForPath "$localThemePath\$themeName\$_"
        if ($testresult -eq $null) {
            Write-Host "$localThemePath\$themeName\$_ does not have a valid binary path saved. If you are creating a new theme, see this file to update: $($MyInvocation.ScriptName)"
            $status = 'Fail'
        }
        if ((Test-Path "$networkThemePath\$themeName\$_") -eq $false) {
            Write-Host "$networkThemePath\$themeName\$_ does not exist. Check the path, ensure the network path is reachable, and that the directories are populated correctly" 
            $status = 'Fail'
        }
        if ((Test-Path "$localThemePath\$themeName\$_") -eq $false) {
            Write-Host "$localThemePath\$themeName\$_ does not exist. The theme folder may not have yet been replicated."
            $status = 'Fail'
        }
        foreach($file in $ThemeFiles) {
            if ((Test-Path "$networkThemePath\$themeName\$file") -eq $false) {
                Write-Host "$networkThemePath\$themeName\$file should exist but does not."
                $status = 'Fail'
            }
        }
    }
    Write-Output $status

    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}

####### Main function called by login script to start everything #######
function Set-Theme {
    Param (
        [String]$themeName,
        [String]$localThemePath = "C:\users\Public\Public Pictures\Themes",
        [String]$networkThemePath = "\\jakartafs01.eap.state.sbu\DesktopThemes$"
    )

    Write-TraceCall "$($MyInvocation.InvocationName)"

    if ((Test-Theme -themeName $themeName) -like 'Success') {
        Set-LockScreenImage -pathToTheme "$localThemePath\$themeName"

        Set-LockScreenSlideshow -PathToSlides "$localThemePath\$themeName\LockScreenSlides"

        $DesktopSlidePath = Choose-BestDesktopSlideDimensions -RootPath "$localThemePath\$themeName\DesktopSlides"
        Set-DesktopSlideshow -PathToSlides $DesktopSlidePath

        Save-LocalStateData
    }
    Write-TraceReturn -msg "$($MyInvocation.InvocationName)"
}