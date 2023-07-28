netstat -ano | Where-Object{$_ -match 'LISTENING'} | 
    ForEach-Object -Begin:{
        New-Variable -Force -Name:'Results' -Value:@()
        New-Variable -Force -Name:'Services' -Value:(
            Get-WmiObject -Class:'Win32_Service' `
                -Filter:"State <> 'Stopped'" `
                -Property:'Name','ProcessId','State','PathName'
        )
        New-Variable -Force -Name:'knownPorts' -Value:@{
            '80'='HTTP';'443'='HTTPS';'135'='RPC';'445'='SMB';'5985'='WinRM'
            '8005'='SCCM'
            '2701'='SMS Remote Control (control)'
            '3389'='Remote Desktop Services'
        }
    } -Process:{
        Set-Variable -Name:'stdOut' -Value:($_.trim())
        If ([String]::IsNullOrEmpty($stdOut) -eq $False) {
        Set-Variable -Name:'parsedOut' -Value:($stdOut -split "\s+")
            If ($parsedOut.count -eq 5) {
                $process = Get-Process -Id:$parsedOut[-1]
                [bool]$isService = $Process.ID -in $Services.ProcessID+4
                If ($isService -eq $True) {
                    $procPath = $Services.where({$_.ProcessID -eq $Process.ID}).PathName
                } Else {
                    $procPath = $Process.path
                }
                [PSCustomObject]@{
                    'Proto'         = $parsedOut[0]
                    'localIP'       = $parsedOut[1].Substring(0,$parsedOut[1].LastIndexOf(':'))
                    'localPort'     = ($parsedOut[1] -split ':')[-1]
                    'localPortName' = $knownPorts[(($parsedOut[1] -split ':')[-1])]
                    'foreignIP'     = $parsedOut[2].Substring(0,$parsedOut[2].LastIndexOf(':'))
                    'foreignPort'   = ($parsedOut[2] -split ':')[-1]
                    'State'         = if($parsedOut[3] -notmatch "\d+"){$parsedOut[3]}else{""}
                    'isService'     = $isService
                    'procName'      = $Process.ProcessName
                    'procID'        = $Process.ID
                    'procPath'      = $procPath
                }
            }
        }
    } |Format-Table -AutoSize
