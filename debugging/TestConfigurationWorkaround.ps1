

$realPowershellPath = 'c:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe'

$args = @("-ConfigurationName microsoft.powershell -NonInteractive -NoProfile -NoLogo -WindowStyle Minimized -File '$PSScriptRoot\TestWinRTall.ps1' -themeName $themeName")
$proc = Start-Process -FilePath $realPowershellPath -ArgumentList $args -PassThru
$proc.WaitForExit(20000)