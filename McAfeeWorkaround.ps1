# Hideous hack to work-around a crash caused by McAfee Host Intrustion Prevention on Windows 10.0.14393. 
# Later versions of Windows seem to work fine.
# McAfee inserts itself into any 'powershell.exe' process.
# By copying powershell.exe to myposh.exe, we bypass McAfee's intrusion and work-around the crash.

Param (
    $script = "Set-Theme.ps1",
    $themeName = "JakartaTest"
)

if ((Get-CimInstance -ClassName Win32_OperatingSystem).Version -like '10.0.14393'){
    $realPowershellPath = 'c:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe'
    $temporaryFakePowershellPath = 'c:\tempdeleteme\myposh.exe'
    $temporaryDeleteme = 'c:\tempdeleteme'
    mkdir $temporaryDeleteme
    Copy-Item $realPowershellPath $temporaryFakePowershellPath

    try {
        
        $args = @("-ConfigurationName microsoft.powershell -NonInteractive -NoProfile -NoLogo -WindowStyle Minimized -File $PSScriptRoot\$script -themeName $themeName")
        $proc = Start-Process -FilePath $temporaryFakePowershellPath -ArgumentList $args -PassThru
        $proc.WaitForExit(20000)
    } finally {
        Remove-Item $temporaryDeleteme -Force -Recurse
    }
} else {
    & $PSScriptRoot\$script -themeName $themeName
}