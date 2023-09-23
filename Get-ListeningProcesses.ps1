# Use a custom class to enforce strict data typing.
Class ListeningPort {
    #Initial Required Values
    [String]$Protocol
    [String]$LocalAddress
    [Int]   $LocalPort
    [int]   $ProcessID
    [String]$RemoteAddress

    # Default Values
    [String]$State = 'Listening'
    
    # Calculated Values
    [String]$localPortName
    [Bool]  $isService = $False
    [String]$ProcessName
    [String]$ProcessPath

    ListeningPort(
        [String]$Protocol,
        [PSCustomObject]$InputObject
    ) {
        $This.Protocol = $Protocol.ToUpper()
        $This.LocalAddress = $InputObject.LocalAddress
        $This.LocalPort = $InputObject.LocalPort
        $This.ProcessID = $InputObject.OwningProcess
    }
}
Class Processes {
    [String]$ProcessName
    [String]$ProcessID
    [System.IO.FileInfo]$ProcessPath

    Processes(
        [CimInstance]$Process
    ){
        $This.ProcessID = $Process.ProcessID

        Switch ($This.ProcessID) {
            4 {
                $This.ProcessName = 'System'
            }
            Default {
                Try {
                    $ProcessString = $Process.ExecutablePath,$Process.CommandLine -join ' '
                    $this.ProcessPath = [Regex]::Match($ProcessString,'^.*?(\b\w:\\.*?.exe\b)') |
                        Select-Object -ExpandProperty:'Groups' |
                        Select-Object -Skip:1 -ExpandProperty:'Value'
                } Catch {
                    Write-Host -Message:'Unable to match.'
                }
                $This.ProcessName = ([System.IO.FileInfo]$Process.Name).BaseName
            }
        }

        
    }
}

Write-Verbose -Message:'Building reusable DCOM CimSession.'
New-Variable -Force -Name:'CIMSession' -Value:(
     New-CimSession -ComputerName:'LOCALHOST' -SessionOption:(
        New-CimSessionOption -Protocol:'DCOM'
    )
)



New-Variable -Force -Name:'listeningPorts' -Value:(New-Object -TypeName:'System.Collections.ArrayList')

Write-Verbose -Message:'Get TCP & UDP listening ports.'
$PortSplat = @{Property = 'LocalAddress','LocalPort','OwningProcess','RemoteAddress'}

Write-Verbose -Message:'Getting TCP listening ports.'
$_tcpPorts = Get-NetTCPConnection -State:'Listen' | Select-Object @PortSplat
$_tcpPorts | & { Process { $Null = $listeningPorts.Add([ListeningPort]::new('TCP',$_)) } }

Write-Verbose -Message:'Getting UDP listening ports.'
$_udpPorts = Get-NetUDPEndpoint | Select-Object @PortSplat
$_udpPorts | & { Process { $Null = $listeningPorts.Add([ListeningPort]::new('UDP',$_)) } }

# region ======= [ Load System Process Information ] =========================================================== #
'','','' | & { Process {}}

# endregion ==== [ Load System Process Information ] =========================================================== #

Function Initialize-Variable {
    Param (
        [Parameter(Mandatory = $True, ParameterSetName = 'Default')]
        [String]$Name,
        [Parameter(Mandatory = $True, ParameterSetName = 'Hashtable')]
        [Hashtable]$Variable
    )

    If ($PSCmdlet.ParameterSetName -eq 'Default') {
        $Value = $Null
    } Else {
        $Name = $Variable.Keys[0]
        $Value = $Variable.Values[0]
    }
    New-Variable -Force -Scope:1 -Name:$Name -Value:$Value

}
New-Variable -Force -Name:'Processes' -Value:(New-Object -TypeName:'System.Collections.ArrayList')
New-Variable -Force -Name:'_CimProcesses' -Value:(
New-Variable -Force -Name:'_GetProcesses' -Value:(

Get-CimInstance -CimSession:$CIMSession -Namespace:'root/cimV2' -ClassName:'Win32_Process' `
    -Property:('ProcessID','ExecutablePath','Name','CommandLine','*') |
        Where-Object -FilterScript:{
            $_.ProcessId -in $listeningPorts.ProcessID
        } |
        & { Process { $Null = $Processes.Add([Processes]::new($_)) } }

$Processes
