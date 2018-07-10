
function Parse-Logfile {
    Param (
        [Parameter (ValueFromPipeline=$true)]
        $logfile
    )

    Process {
        $logfile | ForEach-Object {
            $outObj = [ordered]@{}        
    
            $Filters = @(
                @{Name='Filename';  filter="Transcript started, output file is (.*`.txt)$" ;default=""},
                @{Name='StartTime'; filter="Start time: (.*)" ;default=""},
                @{Name='User';      filter="RunAs User: (.*)" ;default=""},
                @{Name='Machine';   filter="Machine: (.*)" ;default=""},
                @{Name='EndTime';   filter="End time: (.*)$"   ;default=""}
            )

            $contentLines = Get-Content $_

            foreach ($line in $contentLines) {            
                foreach ($filter in $filters) {
                    if ($line -match $filter.filter) {
                        $outObj[$filter.Name] = $Matches[1]
                    }
                }
            }
    
            foreach ($Filter in $filters) {
                if ($outObj.Keys -notcontains $Filter.Name) {
                    $outObj[$filter.Name] = $filter.default
                }
            }
            [pscustomobject]$outObj
        }
    }
}

function Find-LatestFiles {
    Param (
        $logLocation = "\\jakartafs01.eap.state.sbu\ISCProjects\WorkstationInfo\Logs"
    )
    Get-ChildItem $logLocation |Where-Object lastwritetime -gt ([DateTime]::Today).AddDays(-6)
}
