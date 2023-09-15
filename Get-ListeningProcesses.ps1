<#
Proto localIP       localPort localPortName                foreignIP foreignPort State     isService procName                      procID
----- -------       --------- -------------                --------- ----------- -----     --------- --------                      ------
TCP   0.0.0.0       80        HTTP                         0.0.0.0   0           LISTENING      True System                             4

#>

$Ports = @{
    123 = @{ 'svchost.exe' = @{TCP = 'Network Time Protocol (NTP)'; UDP = 'Network Time Protocol (NTP)'}}
    137 = @{ 'svchost.exe' = @{TCP = 'NETBIOS Name Service'; UDP = 'NETBIOS Name Service'}}
    138 = @{ 'svchost.exe' = @{TCP = 'NETBIOS Datagram Service'; UDP = 'NETBIOS Datagram Service'}}
    139 = @{ 'svchost.exe' = @{TCP = 'NETBIOS Session Service'; UDP = 'NETBIOS Session Service'}}
    445 = @{ 'svchost.exe' = @{TCP = 'SMB (Server Message Block)'; UDP = 'SMB (Server Message Block)'}}
    3389 = @{ 'svchost.exe' = @{TCP = 'Microsoft Terminal Server (RDP)'; UDP = 'Microsoft Terminal Server (RDP)'}}
    4500 = @{ 'svchost.exe' = @{UDP = 'IPSec - NAT traversal';}}
    5985 = @{ 'svchost.exe' = @{UDP = 'WinRM 2.0';}}
}

Try {
    $test = $Ports[137]['svchost.exe']['TCP']
} Catch {
    
}

Class ListeningPort {
    [String]$Proto
    [String]$LocalIP
    [Int]   $localPort
    [String]$localPortName
    [String]$State = 'Listening'
    [Bool]  $isService = $False
    [String]$ProcessName
    [String]$ProcessID
    [String]$ProcessPath

    #Constructor
    ListeningPort(
        [String]$Proto,
        [String]$LocalIP,
        [Int]   $localPort,
        [String]$ProcessID,
        [Array] $Processes,
        [Array] $Services
    ) {
    
        Write-Verbose -Message:'Loading process information.'
        $Process = $Processes.Where({$_.ProcessId -eq $ProcessId})

        # Error handling for process.
        If ([string]::IsNullOrEmpty($Process) -eq -$true) {
            Throw 'Unable to locate a matching process; exiting.'
            exit 1
        }

        Write-Verbose -Message:'Finding services with matching ProcessID.'
        $Service = $Services.Where({$_.ProcessId -eq $ProcessId})
    
        # If no services were found, then this is empty so invert.
        $this.isService = -NOT [String]::IsNullOrEmpty($Service)
    
        # Some mundane settings.
        $this.Proto = $Proto
        $this.LocalIP = $LocalIP
        $This.localPort = $LocalPort
        $this.ProcessID = $ProcessID
        $this.ProcessName = $process.Name

        If ($This.isService -eq $True) {
            If ([String]::IsNullOrEmpty($This.ProcessPath) -and $This.isService -eq $True) {
                Try {
                $this.ProcessPath = [Regex]::Match($Service.PathName,'^.*?(\b\w:\\.*?.exe\b)') |
                    Select-Object -ExpandProperty:'Groups' |
                    Select-Object -Skip:1 -ExpandProperty:'Value'
                } Catch {
                    Write-Verbose -Message:'Failed to find path match.'
                }
            }
        } ElseIf ($Null -NE $process.ExecutablePath) {
            Try {
                $this.ProcessPath = ([System.IO.FileInfo]$process.ExecutablePath).FullName
            } Catch {
                Write-Warning -Message:'Unable to convert ExecutablePath to fullname.'
            }
        } Else {
            $this.ProcessPath = Get-Process -Id:$ProcessID | Select-Object -ExpandProperty:'Path'
        }
    }
}

New-Variable -Force -Name:'Results' -Value:(New-Object -TypeName:'System.Collections.ArrayList')
New-Variable -Force -Name:'CimSession' -Value:(New-CimSession -ComputerName:'LOCALHOST' `
    -SessionOption:(New-CimSessionOption -Protocol:'DCOM'))

New-Variable -Force -Name:'Processes' -Value:(Get-CimInstance -CimSession:$cimSession `
    -Namespace:'root/cimV2' -ClassName:'Win32_Process' -Property:('ProcessID',
    'ExecutablePath','Name','*')
)

New-Variable -Force -Name:'Services' -Value:(Get-CimInstance -CimSession:$cimSession `
    -Namespace:'root/cimV2' -ClassName:'Win32_Service' -Filter:"(State <> 'Stopped')" `
    -Property:('Name','ProcessId','State','PathName')
)

# Query TCP Ports
Get-NetTCPConnection -State:'Listen' | ForEach-Object -Process:{
    $Null = $Results.Add([ListeningPort]::NEW('TCP',$_.LocalAddress,$_.LocalPort,$_.OwningProcess,$Processes,$Services))
}

Get-NetUDPEndpoint | ForEach-Object -Process:{
    $Null = $Results.Add([ListeningPort]::NEW('UDP',$_.LocalAddress,$_.LocalPort,$_.OwningProcess,$Processes,$Services))
}
