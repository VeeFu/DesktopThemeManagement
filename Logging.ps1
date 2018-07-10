Param (    
    $LogLevel = @('Trace','Debug'),
    $LogFile = "\\jakartafs01.eap.state.sbu\workstationInfo$\Logs\$env:computername.trace.log"
)

if ($Debug -eq $true) {
    Start-Transcript -Path '\\jakartafs01.eap.state.sbu\workstationInfo$\Logs' -Append
}

function Write-Trace {
    Param ($msg)
    if ($LogLevel -contains 'Trace') { Write-Log -level 'Trace' -msg $msg }
}

function Write-TraceCall {
    Param($msg)
    Write-Trace -msg "Call $msg"
}

function Write-TraceReturn {
    Param($msg)
    Write-Trace -msg "Return $msg"
}

function Write-Debug {
    Param ($msg)
    if ($LogLevel -contains 'Debug') {Write-Log -level 'Debug' -msg $msg}
}

function Write-Log {
    Param ($level, $msg)
    "$([DateTime]::Now.ToString("s"))  $level  $msg" | Out-File $LogFile -Append
}
