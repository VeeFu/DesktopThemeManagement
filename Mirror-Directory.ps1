#copy images if necessary
Param (    
    $RemoteSource = "\\jakartafs01.eap.state.sbu\$",
    $LocalDestination = "C:\users\Public\Pictures\Themes"
)
if ((Test-Path $LocalDestination) -eq $false) {
    mkdir $LocalDestination
}

robocopy /MIR /NP $RemoteSource $LocalDestination | Out-File "c:\users\public\Mirror-Directory.log"