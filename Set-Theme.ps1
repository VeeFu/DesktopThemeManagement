Param ($themeName)

Start-Transcript -OutputDirectory '\\jakartafs01.eap.state.sbu\WorkstationInfo$\Logs' -Append

. $PSScriptRoot\ThemeManagement.ps1

Set-Theme $themeName
