$transfile = "C:\temps\scripts\DesktopTheme\debugging\WinRTTrans.log"
$logfile = "C:\temps\scripts\DesktopTheme\debugging\WinRTLog.log"

Start-Transcript -Path $transfile

Add-Type -AssemblyName System.Runtime.WindowsRuntime
Add-Type -AssemblyName System.Windows.Forms
    [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime] | Out-Null
    [Windows.System.UserProfile.LockScreen,Windows.System.UserProfile,ContentType=WindowsRuntime] | Out-Null

$asTaskGeneric = (
    [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object { 
            $_.Name -eq 'AsTask' -and 
            $_.GetParameters().Count -eq 1 -and 
            $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' 
        }
)[0]

Function Await($WinRtTask, $ResultType) {

    "Calling $($MyInvocation.InvocationName)" | Out-File -Append -FilePath $logfile
    "  Calling MakeGenericMethod" | Out-File -Append -FilePath $logfile
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    
    "  Calling Invoke" | Out-File -Append -FilePath $logfile
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    
    "  Calling Wait" | Out-File -Append -FilePath $logfile
    $netTask.Wait(-1) | Out-Null
    
    $result = $netTask.Result
    "  Getting Result: $($result)" | Out-File -Append -FilePath $logfile
    
    Write-Output $result
    "Returning from $($MyInvocation.InvocationName)" | Out-File -Append -FilePath $logfile
}

function Set-LockScreenImage {
    Param ($newLockScreenPath = 'C:\users\public\Public Pictures\Themes\CatsTest\LockScreen1680x1050.jpg')

    "Calling $($MyInvocation.InvocationName)" | Out-File -Append -FilePath $logfile
        
    if (Test-Path $newLockScreenPath) {
        
        " Calling GetFileFromPathAsync" | Out-File -FilePath $logfile -Append
        $getFileFromPathAsyncTask = [Windows.Storage.StorageFile]::GetFileFromPathAsync($newLockScreenPath)
        
        $newLockScreen = Await -WinRtTask $getFileFromPathAsyncTask -ResultType ([Windows.Storage.StorageFile])

        "  Calling SetImageFileAsync" | Out-File -FilePath $logfile -Append
        try {
            $setImageFileAsyncTask = [Windows.System.UserProfile.LockScreen]::SetImageFileAsync($newLockScreen)
        } catch {
            "  Caught a thing $_" | Out-File -FilePath $logfile -Append
        }
    }
    
    "Returning from $($MyInvocation.InvocationName)" | Out-File -Append -FilePath $logfile
}

Set-LockScreenImage