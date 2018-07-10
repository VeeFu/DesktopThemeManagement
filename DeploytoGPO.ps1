Param($chosenGPO = 'Set Animals Theme')

$deplomentLocation = @{
    'Set Mission Goals Theme' = '\\eap.state.sbu\sysvol\eap.state.sbu\Policies\{3A94AF8A-6F19-4B1A-BB45-9479322D69C5}\User\Scripts\Logon'
    'Set Animals Theme' =       '\\eap.state.sbu\sysvol\eap.state.sbu\Policies\{78EE19A7-8277-492B-BC2A-3E6ED6AEC41A}\User\Scripts\Logon'
}

$sourceLocation = "$PSScriptRoot"

$filesToDeploy = @(
    'Set-Theme.ps1',
    'ThemeManagement.ps1',
    'DataCollection.xml',
    'PIDLStringCollection.xml'
    'McAfeeWorkaround.ps1'
    'Logging.ps1'
    'PIDLStringCache.ps1'    
)

$filesToDeploy | ForEach-Object {
    Copy-Item "$sourceLocation\$_" $deplomentLocation[$chosenGPO]
}