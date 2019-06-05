function Invoke-AsBuiltReport.Cisco.UcsManager {
    <#
    .SYNOPSIS  
        PowerShell script to document the configuration of Cisco UCS infrastucture in Word/HTML/XML/Text formats
    .DESCRIPTION
        Documents the configuration of Cisco UCS infrastucture in Word/HTML/XML/Text formats using PScribo.
    .NOTES
        Version:        0.2.0
        Author:         Tim Carman
        Twitter:        @tpcarman
        Github:         tpcarman
        Credits:        Brandon Beck (robbeck@cisco.com) - Cisco UCS Health Check Script
                        Martijn Smit (@smitmartijn) - Cisco UCS Inventory Script
                        Iain Brighton (@iainbrighton) - PScribo module
    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.Cisco.UcsManager
    #>

    param (
        [String[]] $Target,
        [PSCredential] $Credential,
        [String]$StylePath
    )

    # Import JSON Configuration for Section and InfoLevel
    $InfoLevel = $ReportConfig.InfoLevel
    $Section = $ReportConfig.Section
    
    #region Configuration Settings
    # If custom style not set, use default style
    if (!$StylePath) {
        & "$PSScriptRoot\..\..\AsBuiltReport.Cisco.UcsManager.Style.ps1"
    }

    foreach ($UCS in $Target) {
        # Connect to Cisco UCS domain using supplied credentials
        $UCSM = @()       
        try {
            $UCSM = Connect-Ucs -Name $UCS -Credential $Credential -ErrorAction Stop
        } catch {
            Write-Error $_
        }
        #endregion Configuration Settings

        #region Script Body
        #---------------------------------------------------------------------------------------------#
        #                                       SCRIPT BODY                                           #
        #---------------------------------------------------------------------------------------------#

        # Generate report if connection to UCS Manager is successful
        if ($UCSM) {
            #region Collect UCS Info
            $UcsLicense = Get-UcsLicense -Ucs $UCSM
            $UcsLanCloud = Get-UcsLanCloud -Ucs $UCSM
            $UcsSanCloud = Get-UcsSanCloud -Ucs $UCSM
            $UcsSystem = Get-UcsStatus -Ucs $UCSM
            $UcsFi = Get-UcsNetworkElement -Ucs $UCSM | Sort-Object Id
            $UcsBlade = Get-UcsBlade -Ucs $UCSM
            $UcsRackUnit = Get-UcsRackUnit -Ucs $UCSM | Sort-Object Id
            #endregion Collect UCS Info

            Section -Style Heading1 -Name "$($UcsSystem.Name)" {
                #region System Section
                if ($Section.System) {
                    if ($UcsSystem) {
                        Section -Style Heading2 -Name 'System' {
                            #region Cluster Summary Table
                            $ClusterStatus = [PSCustomObject]@{
                                'Name' = $UcsSystem.Name 
                                'Virtual IP Address' = $UcsSystem.VirtualIpv4Address
                                'HA Configuration' = $UcsSystem.HaConfiguration 
                                'HA Ready' = $UcsSystem.HaReady 
                                'Ethernet State' = $UcsSystem.EthernetState
                                'Backup Policy' = (Get-UcsMgmtBackupPolicy -Ucs $UCSM).AdminState
                                'Config Policy' = (Get-UcsMgmtCfgExportPolicy -Ucs $UCSM).AdminState
                                'Call Home State' = (Get-UcsCallhome -Ucs $UCSM).AdminState
                                'UCSM System Version' = (Get-UcsMgmtController -Ucs $UCSM -Subject system | Get-UcsFirmwareRunning).Version | Select-Object -Last 1
                            }
                            if ($Healthcheck.Cluster.HAReady) {
                                $ClusterStatus | Where-Object { $_.'HA Ready' -ne 'yes' } | Set-Style -Style Critical -Property 'HA Ready'
                            }
                            $ClusterStatus | Table -Name 'Cluster Status' -List -ColumnWidths 50, 50 
                            #endregion Cluster Summary Table
                        
                            BlankLine

                            #region Fabric Interconnect A Summary Table
                            $UcsStatusFiA = [PSCustomObject]@{
                                'Fabric Interconnect A Role' = $UcsSystem.FiALeadership 
                                'Fabric Interconnect A IP Address' = $UcsSystem.FiAOobIpv4Address 
                                'Fabric Interconnect A Subnet Mask' = $UcsSystem.FiAOobIpv4SubnetMask 
                                'Fabric Interconnect A Default Gateway' = $UcsSystem.FiAOobIpv4DefaultGateway 
                                'Fabric Interconnect A State' = $UcsSystem.FiAManagementServicesState
                            }
                            if ($Healthcheck.FI.State) {
                                $ClusterStatus | Where-Object { $_.'Fabric Interconnect A State' -ne 'up' } | Set-Style -Style Critical -Property 'Fabric Interconnect A State'
                            }
                            $UcsStatusFiA | Table -Name 'Fabric Interconnect A Information' -List -ColumnWidths 50, 50 
                            #endregion Fabric Interconnect A Summary Table

                            BlankLine

                            #region Fabric Interconnect B Summary Table
                            $UcsStatusFiB = [PSCustomObject]@{
                                'Fabric Interconnect B Role' = $UcsSystem.FiBLeadership 
                                'Fabric Interconnect B IP Address' = $UcsSystem.FiBOobIpv4Address 
                                'Fabric Interconnect B Subnet Mask' = $UcsSystem.FiBOobIpv4SubnetMask
                                'Fabric Interconnect B Default Gateway' = $UcsSystem.FiBOobIpv4DefaultGateway 
                                'Fabric Interconnect B State' = $UcsSystem.FiBManagementServicesState
                            }
                            if ($Healthcheck.FI.State) {
                                $ClusterStatus | Where-Object { $_.'Fabric Interconnect B State' -ne 'up' } | Set-Style -Style Critical -Property 'Fabric Interconnect B State'
                            }
                            $UcsStatusFiB | Table -Name 'Fabric Interconnect B Information' -List -ColumnWidths 50, 50 
                            #endregion Fabric Interconnect B Summary Table
                        
                            BlankLine

                            #region Fault Summary Table
                            $UcsFault = Get-UcsFault -Ucs $UCSM
                            if ($UcsFault -and $Options.ShowFaultSummary) {
                                $UcsFaults = [PSCustomObject]@{
                                    'Critical Faults' = ($UcsFault | Where-Object { $_.Severity -eq 'critical' }).Count
                                    'Major Faults' = ($UcsFault | Where-Object { $_.Severity -eq 'major' }).Count
                                    'Minor Faults' = ($UcsFault | Where-Object { $_.Severity -eq 'minor' }).Count
                                    'Warnings' = ($UcsFault | Where-Object { $_.Severity -eq 'warning' }).Count
                                }
                                $UcsFaults | Table -Name 'UCS Faults' -ColumnWidths 25, 25, 25, 25
                            }
                            #endregion Fault Summary Table
                        }
                    }
                }
                #endregion System Section

                #region Equipment Section
                if ($Section.Equipment) {
                    Section -Style Heading2 -Name 'Equipment' {
                        #region Faric Interconnect Section
                        if ($InfoLevel.Equipment.FabricInterconnects -ge 1) {
                            Section -Style Heading3 -Name 'Fabric Interconnects' {
                                #region Faric Interconnect Summary Table
                                $UcsFiSummary = foreach ($Fi in $UcsFi) {
                                    $FiBoot = Get-UcsMgmtController -Ucs $UCSM -Dn "$($fi.Dn)/mgmt" | Get-UcsFirmwareBootDefinition | Get-UcsFirmwareBootUnit -Filter 'Type -ieq system -or Type -ieq kernel' | Select-Object Type, Version
                                    [PSCustomObject]@{
                                        'Fabric' = $Fi.Id
                                        'Cluster Role' = Switch ($Fi.Id) {
                                            'A' { $UcsSystem.FiALeadership }
                                            'B' { $UcsSystem.FiBLeadership }
                                        }
                                        'Model' = $Fi.Model
                                        'Serial' = $Fi.Serial
                                        'System' = ($FiBoot | Where-Object { $_.Type -eq "system" }).Version
                                        'Kernel' = ($FiBoot | Where-Object { $_.Type -eq "kernel" }).Version
                                        'Ports Used' = ((Get-UcsLicense -Ucs $UCSM -Scope $Fi.Id).UsedQuant | Measure-Object -Sum).Sum
                                        'Ports Licensed' = ((Get-UcsLicense -Ucs $UCSM -Scope $Fi.Id).AbsQuant | Measure-Object -Sum).Sum
                                        'Ethernet Mode' = $UcsLanCloud.Mode
                                        'FC Mode' = $UcsSanCloud.Mode
                                        'Status' = $Fi.Operability
                                        'Thermal' = $Fi.Thermal
                                    }
                                }
                                $UcsFiSummary | Sort-Object 'Fabric' | Table -Name 'Fabric Interconnects'
                                #endregion Faric Interconnect Summary Table

                                #region Faric Interconnect Detailed Section
                                if ($InfoLevel.Equipment.FabricInterconnects -ge 2) {
                                    foreach ($Fi in $UcsFi) {
                                        Section -Style Heading3 "Fabric Interconnect $($Fi.Id)" {
                                            $FiBoot = Get-UcsMgmtController -Ucs $UCSM -Dn "$($Fi.Dn)/mgmt" | Get-UcsFirmwareBootDefinition | Get-UcsFirmwareBootUnit -Filter 'Type -ieq system -or Type -ieq kernel' | Select-Object Type, Version
                                            #region Fabric Interconnect Detailed Table
                                            $UcsFiDetailed = [PSCustomObject]@{
                                                'Fabric' = $Fi.Id
                                                'Cluster Role' = Switch ($Fi.Id) {
                                                    'A' { $UcsSystem.FiALeadership }
                                                    'B' { $UcsSystem.FiBLeadership }
                                                }
                                                'IP Address' = $Fi.OobIfIp
                                                'Subnet Mask' = $Fi.OobIfMask
                                                'Default Gateway' = $Fi.OobIfGw
                                                'MAC Address' = $Fi.OobIfMac
                                                'Model' = $Fi.Model
                                                'Serial' = $Fi.Serial
                                                'Total Memory (GB)' = [math]::Round(($Fi.TotalMemory / 1024), 3)
                                                'System' = ($FiBoot | Where-Object { $_.Type -eq "system" }).Version
                                                'Kernel' = ($FiBoot | Where-Object { $_.Type -eq "kernel" }).Version
                                                'State' = Switch ($Fi.Id) {
                                                    'A' { $UcsSystem.FiAManagementServicesState }
                                                    'B' { $UcsSystem.FiBManagementServicesState }
                                                }
                                                'Ports Used' = ((Get-UcsLicense -Ucs $UCSM -Scope $Fi.Id).UsedQuant | Measure-Object -Sum).Sum
                                                'Ports Licensed' = ((Get-UcsLicense -Ucs $UCSM -Scope $fi.Id).AbsQuant | Measure-Object -Sum).Sum
                                                'Ethernet Mode' = $UcsLanCloud.Mode
                                                'FC Mode' = $UcsSanCloud.Mode
                                                'Status' = $Fi.Operability
                                                'Thermal' = $Fi.Thermal
                                                'Admin Evac Mode' = Switch ($Fi.AdminEvacState) {
                                                    'fill' { 'off' }
                                                    default { $Fi.AdminEvacState }
                                                }
                                                'Oper Evac Mode' = Switch ($Fi.OperEvacState) {
                                                    'fill' { 'off' }
                                                    default { $Fi.OperEvacState }
                                                }
                                            }
                                            $UcsFiDetailed | Sort-Object 'Fabric' | Table -List -Name "Fabric Interconnect $($Fi.Id)"
                                            #endregion Fabric Interconnect Detailed Table

                                            #region Fabric Interconnect Fixed Module Section
                                            $UcsFiModule = $Fi | Get-UcsFiModule -Ucs $UCSM
                                            Section -Style Heading4 -Name 'Fixed Module' {
                                                $UcsFiFixedModule = [PSCustomObject]@{
                                                    'ID' = $UcsFiModule.Id
                                                    'Model' = $UcsFiModule.Model
                                                    'Serial' = $UcsFiModule.Serial
                                                    'Description' = $UcsFiModule.Descr
                                                    'Max Number of Ports' = $UcsFiModule.NumPorts
                                                    'Status' = $UcsFiModule.OperState
                                                }
                                                $UcsFiFixedModule | Sort-Object 'ID' | Table -Name 'Fixed Modules' 
                                            }
                                            #endregion Fabric Interconnect Fixed Module Section

                                            #--- Sort Expression to filter port id to be just the numerical port number and sort ascending ---#
                                            $sortExpr = { if ($_.Dn -match "(?=port[-]).*") { ($matches[0] -replace ".*(?<=[-])", '') -as [int] } }
                                            #--- Get Fabric Port Configuration and sort by port id using the above sort expression ---#
                                    
                                            #region Fabric Interconnect Ethernet Ports Section
                                            $UcsEthernetPorts = Get-UcsFabricPort -Ucs $UCSM -SwitchId "$($fi.Id)" | Sort-Object $sortExpr 
                                            if ($UcsEthernetPorts) {
                                                Section -Style Heading4 -Name 'Ethernet Ports' {
                                                    $UcsFiEthernetPorts = foreach ($UcsEthernetPort in $UcsEthernetPorts) {
                                                        [PSCustomObject]@{
                                                            'Slot' = $UcsEthernetPort.SlotId
                                                            'Port ID' = $UcsEthernetPort.PortId
                                                            'MAC' = $UcsEthernetPort.Mac
                                                            'State' = $UcsEthernetPort.AdminState
                                                            'If Role' = $UcsEthernetPort.IfRole
                                                            'If Type' = $UcsEthernetPort.IfType
                                                            'Status' = $UcsEthernetPort.OperState
                                                            'Peer' = $UcsEthernetPort.PeerDn
                                                        } 
                                                    }
                                                    $UcsFiEthernetPorts | Sort-Object 'Slot', 'Port ID' | Table -Name "Fabric Interconnect Ethernet Ports"
                                                }
                                            }
                                            #endregion Fabric Interconnect Ethernet Ports Section
                                    
                                            #region Fabric Interconnect FC Uplink Ports Section
                                            $UcsFcUplinkPorts = Get-UcsFiFcPort -Ucs $UCSM -SwitchId "$($fi.Id)" -IfRole 'network'
                                            if ($UcsFcUplinkPorts) {
                                                Section -Style Heading4 -Name 'FC Uplink Ports' {
                                                    $UcsFiFcUplinkPorts = foreach ($UcsFcUplinkPort in $UcsFcUplinkPorts) {
                                                        [PSCustomObject]@{
                                                            'Slot' = $UcsFcUplinkPort.SlotId
                                                            'Port ID' = $UcsFcUplinkPort.PortId
                                                            'WWPN' = $UcsFcUplinkPort.WWN
                                                            'State' = $UcsFcUplinkPort.AdminState
                                                            'If Role' = $UcsFcUplinkPort.IfRole
                                                            'If Type' = $UcsFcUplinkPort.IfType
                                                            'Status' = $UcsFcUplinkPort.OperState
                                                        }
                                                    }
                                                    $UcsFiFcUplinkPorts | Sort-Object 'Slot', 'Port ID' | Table -Name "Fabric Interconnect FC Uplink Ports"
                                                }
                                            }
                                            #endregion Fabric Interconnect FC Uplink Ports Section

                                            #region Fabric Interconnect Fans Section
                                            $UcsFiFans = Get-UcsFan -NetworkElement $Fi -Ucs $UCSM 
                                            if ($UcsFiFans) {
                                                Section -Style Heading4 -Name 'Fans' {
                                                    $FiFans = foreach ($UcsFiFan in $UcsFiFans) {
                                                        [PSCustomObject]@{
                                                            'Name' = "Fan Module $($UcsFiFan.Module)"
                                                            'Fan' = $UcsFiFan.Id
                                                            'Model' = $UcsFiFan.Model
                                                            'Power' = $UcsFiFan.Power
                                                            'Thermal' = $UcsFiFan.Thermal
                                                            'Presence' = $UcsFiFan.Presence
                                                            'Operability' = $UcsFiFan.OperState
                                                        }
                                                    }
                                                    $FiFans | Sort-Object 'Name', 'Fan' | Table -Name "Fabric Interconnect Fans"
                                                }
                                            }
                                            #endregion Fabric Interconnect Fans Section

                                            #region Fabric Interconnect PSUs Section
                                            $UcsFiPSUs = Get-UcsPsu -NetworkElement $Fi -Ucs $UCSM 
                                            if ($UcsFiPSUs) {
                                                Section -Style Heading4 -Name 'PSUs' {
                                                    $FiPSUs = foreach ($UcsFiPSU in $UcsFiPSUs) {
                                                        [PSCustomObject]@{
                                                            'Name' = "PSU $($UcsFiPSU.Id)"
                                                            'ID' = $UcsFiPSU.Id
                                                            'Model' = $UcsFiPSU.Model
                                                            'Power' = $UcsFiPSU.Power
                                                            'Voltage' = $UcsFiPSU.Voltage
                                                            'Performance' = $UcsFiPSU.Perf
                                                            'Thermal' = $UcsFiPSU.Thermal
                                                            'Presence' = $UcsFiPSU.Presence
                                                            'Operability' = $UcsFiPSU.OperState
                                                        }
                                                    }
                                                    $FiPSUs | Sort-Object 'Id', 'Name' | Table -Name "Fabric Interconnect PSUs"
                                                }
                                            }
                                            #endregion Fabric Interconnect PSUs Section
                                        }
                                    }
                                }
                                #endregion Faric Interconnect Detailed Section
                            }
                        }
                        #endregion Faric Interconnect Section

                        #region UCS Chassis Section
                        $UcsChassis = Get-UcsChassis -Ucs $UCSM
                        if ($UcsChassis) {
                            if ($InfoLevel.Equipment.Chassis -ge 1) {
                                Section -Style Heading3 -Name 'Chassis' {
                                    #region UCS Chassis Summary Table
                                    $UcsChassisSummary = ForEach ($Chassis in $UcsChassis) {
                                        $ChassisSlotCount = 0
                                        $ChassisSlotCount = (Get-UcsBlade -Chassis $Chassis).Count
                                        $ChassisPsuCount = 0
                                        $ChassisPsuCount = (Get-UcsPsu -Chassis $Chassis).Count
                                        [PSCustomObject]@{
                                            'Chassis ID' = $Chassis.Id 
                                            'Model' = $Chassis.Model
                                            'Serial' = $Chassis.Serial
                                            'Slots Used' = $ChassisSlotCount
                                            'Slots Available' = (8 - $ChassisSlotCount)
                                            'PSUs' = $ChassisPsuCount
                                            'Power' = $Chassis.Power
                                            'Power Redundancy' = (Get-UcsComputePsuControl -Chassis $Chassis).Redundancy
                                            'Thermal' = $Chassis.Thermal
                                            'State' = $Chassis.AdminState 
                                            'Status' = $Chassis.OperState
                                            'Operability' = $Chassis.Operability
                                        }
                                    }
                                    $UcsChassisSummary | Sort-Object 'Chassis ID' | Table -Name 'Chassis Summary'
                                    #endregion UCS Chassis Summary Table

                                    #region UCS Chassis Detailed Section
                                    if ($InfoLevel.Equipment.Chassis -ge 2) {
                                        ForEach ($Chassis in $UcsChassis) {
                                            Section -Style Heading3 -Name "$($Chassis.Rn)" {
                                                #region UCS Chassis Detailed Table
                                                $ChassisSlotCount = 0
                                                $ChassisSlotCount = (Get-UcsBlade -Chassis $Chassis).Count
                                                $ChassisPsuCount = 0
                                                $ChassisPsuCount = (Get-UcsPsu -Chassis $Chassis).Count
                                                $ChassisDetailed = [PSCustomObject]@{
                                                    'Chassis ID' = $Chassis.Id 
                                                    'Model' = $Chassis.Model
                                                    'Serial' = $Chassis.Serial
                                                    'Slots Used' = $ChassisSlotCount
                                                    'Slots Available' = (8 - $ChassisSlotCount)
                                                    'PSUs' = $ChassisPsuCount
                                                    'Power' = $Chassis.Power
                                                    'Power Redundancy' = (Get-UcsComputePsuControl -Chassis $Chassis).Redundancy
                                                    'Thermal' = $Chassis.Thermal
                                                    'License State' = $Chassis.LicState
                                                    'State' = $Chassis.AdminState 
                                                    'Status' = $Chassis.OperState
                                                    'Operability' = $Chassis.Operability
                                                }
                                                $ChassisDetailed | Table -List -Name "$($Chassis.Rn) Detailed Information" -ColumnWidths 50, 50
                                                #endregion UCS Chassis Detailed Table

                                                #region UCS Chassis Fan Section
                                                $UcsChassisFans = Get-UcsFanModule -Chassis $Chassis -Ucs $UCSM
                                                if ($UcsChassisFans) {
                                                    Section -Style Heading4 -Name 'Fans' {
                                                        $ChassisFans = foreach ($UcsChassisFan in $UcsChassisFans) {
                                                            [PSCustomObject]@{
                                                                'Name' = "Fan Module $($UcsChassisFan.Id)"
                                                                'ID' = $UcsChassisFan.Id
                                                                'Power' = $UcsChassisFan.Power
                                                                'Voltage' = $UcsChassisFan.Voltage
                                                                'Performance' = $UcsChassisFan.Perf
                                                                'Thermal' = $UcsChassisFan.Thermal
                                                                'Presence' = $UcsChassisFan.Presence
                                                                'Operability' = $UcsChassisFan.OperState
                                                            }
                                                        }
                                                        $ChassisFans | Sort-Object 'Id', 'Name' | Table -Name "$($Chassis.Rn) Fans"
                                                    }
                                                }
                                                #endregion UCS Chassis Fan Section

                                                #region UCS Chassis IOM Section
                                                $UcsChassisIoms = Get-UcsIom -Chassis $Chassis -Ucs $UCSM
                                                if ($UcsChassisIoms) {
                                                    Section -Style Heading4 -Name 'IO Modules' {
                                                        #region Chassis IOM Inventory Table
                                                        $UcsChassisIom = foreach ($Iom in $UcsChassisIoms) {
                                                            [PSCustomObject]@{
                                                                'Name' = "IO Module $($Iom.Id)"
                                                                'Chassis ID' = "$($Iom.ChassisId)"
                                                                'Fabric' = $Iom.SwitchId
                                                                'Side' = $Iom.Side
                                                                'Model' = $Iom.Model
                                                                'Serial' = $Iom.Serial
                                                                'Discovery' = $Iom.Discovery
                                                                'Config State' = $Iom.ConfigState 
                                                                'Operability' = $Iom.OperState
                                                                'Thermal' = $Iom.Thermal
                                                                'Presence' = $Iom.Presence
                                                                'Running Firmware' = (Get-UcsMgmtController -Ucs $UCSM -Dn "$($Iom.Dn)/mgmt" | Get-UcsFirmwareRunning -Deployment system | Select-Object Version).Version
                                                                'Backup Firmware' = (Get-UcsMgmtController -Ucs $UCSM -Dn "$($iom.Dn)/mgmt" | Get-UcsFirmwareUpdatable | Select-Object Version).Version
                                                            }
                                                        }
                                                        $UcsChassisIom | Sort-Object 'Chassis', 'Name' | Table -Name 'IOM Inventory'
                                                        #endregion Chassis IOM Inventory Table

                                                        if ($InfoLevel.Equipment.Chassis -ge 3) {
                                                            #region Chassis IOM Fabric Ports Section
                                                            $FabricPorts = Get-UcsEtherSwitchIntFIo -Ucs $UCSM | Where-Object { $_.ChassisId -eq "$($Iom.ChassisId)" }
                                                            if ($FabricPorts) {   
                                                                Section -Style Heading4 -Name 'Fabric Ports' {
                                                                    $IomFabricPorts = foreach ($FabricPort in $FabricPorts) {
                                                                        [PSCustomObject]@{    
                                                                            'Name' = 'Fabric Port ' + $FabricPort.SlotId + '/' + $FabricPort.PortId
                                                                            'Oper State' = $FabricPort.OperState
                                                                            'Port Channel' = $FabricPort.EpDn
                                                                            'Peer Slot ID' = $FabricPort.PeerSlotId
                                                                            'Peer Port ID' = $FabricPort.PeerPortId
                                                                            'Fabric' = $FabricPort.SwitchId
                                                                            'Peer' = $FabricPort.PeerDn
                                                                        }
                                                                    }
                                                                    $IomFabricPorts | Table -Name 'IOM Fabric Ports'
                                                                }
                                                            }
                                                            #endregion Chassis IOM Fabric Ports Section
                
                                                            #region Chassis IOM Backplane Ports Section
                                                            $BackplanePorts = Get-UcsEtherServerIntFIo -Ucs $UCSM | Where-Object { $_.ChassisId -eq "$($Iom.ChassisId)" } # -and $_.SwitchId -eq "$($Iom.SwitchId)"}
                                                            $BackplanePorts = $BackplanePorts | Sort-Object { ($_.SlotId) -as [int] }, { ($_.PortId) -as [int] }
                                                            if ($BackplanePorts) {
                                                                Section -Style Heading4 -Name 'Backplane Ports' {
                                                                    $IomBackplanePorts = foreach ($BackplanePort in $BackplanePorts) {
                                                                        [PSCustomObject]@{
                                                                            'Name' = 'Backplane Port ' + $BackplanePort.SlotId + '/' + $BackplanePort.PortId
                                                                            'Oper State' = $BackplanePort.OperState
                                                                            'Port Channel' = $BackplanePort.EpDn
                                                                            'Fabric' = $BackplanePort.SwitchId
                                                                            'Peer' = $BackplanePort.PeerDn
                                                                        }
                                                                    }
                                                                    $IomBackplanePorts | Table -Name 'IOM Backplane Ports'
                                                                }
                                                            }
                                                            #endregion Chassis IOM Backplane Ports Section
                                                        }
                                                    }
                                                }
                                                #endregion UCS Chassis IOM Section

                                                #region UCS Chassis PSU Section
                                                $UcsChassisPSUs = Get-UcsPsu -Chassis $Chassis -Ucs $UCSM
                                                if ($UcsChassisPSUs) {
                                                    Section -Style Heading4 -Name 'PSUs' {
                                                        $ChassisPSUs = foreach ($UcsChassisPSU in $UcsChassisPSUs) {
                                                            [PSCustomObject]@{
                                                                'Name' = "PSU $($UcsChassisPSU.Id)"
                                                                'ID' = $UcsChassisPSU.Id
                                                                'Model' = $UcsChassisPSU.Model
                                                                'Power' = $UcsChassisPSU.Power
                                                                'Voltage' = $UcsChassisPSU.Voltage
                                                                'Performance' = $UcsChassisPSU.Perf
                                                                'Thermal' = $UcsChassisPSU.Thermal
                                                                'Presence' = $UcsChassisPSU.Presence
                                                                'Operability' = $UcsChassisPSU.OperState
                                                            }
                                                        }
                                                        $ChassisPSUs | Sort-Object 'Id', 'Name' | Table -Name "$($Chassis.Rn) PSUs"
                                                    }
                                                }
                                                #endregion UCS Chassis PSU Section

                                                #region UCS Chassis Blades Section
                                                $UcsChassisBlades = Get-UcsBlade -Chassis $Chassis -Ucs $UCSM
                                                if ($UcsChassisBlades) {
                                                    Section -Style Heading4 -Name 'Servers' {
                                                        $UcsChassisBlade = foreach ($Blade in $UcsChassisBlades) {
                                                            [PSCustomObject]@{
                                                                'Name' = "Server $($Blade.SlotId)"
                                                                'Chassis ID' = $Blade.ChassisID 
                                                                'Model' = $Blade.Model
                                                                'Serial' = $Blade.Serial
                                                                'CPUs' = $Blade.NumOfCpus
                                                                'Cores' = $Blade.NumOfCores
                                                                'Threads' = $Blade.NumOfThreads 
                                                                'Memory GB' = $Blade.AvailableMemory / 1024
                                                                'Adapters' = $Blade.NumOfAdaptors 
                                                                'NICs' = $Blade.NumOfEthHostIfs
                                                                'HBAs' = $Blade.NumOfFcHostIfs
                                                                'Power State' = $Blade.OperPower
                                                                'Status' = $Blade.OperState
                                                                'Assoc State' = $Blade.Association
                                                                'Operability' = $Blade.Operability 
                                                            } 
                                                        }
                                                        $UcsChassisBlade | Sort-Object 'Chassis ID', 'Name' | Table -Name "$($Chassis.Rn) Servers"
                                                    }
                                                }
                                                #endregion UCS Chassis Blades Section
                                            }
                                        }
                                    }
                                    #endregion UCS Chassis Detailed Section
                                }
                            }
                        }
                        #endregion UCS Chassis Section

                        #region UCS Rack-Mounts Section
                        #TODO: Rack Mounts
                        if ($UcsRackUnit) {
                            if ($InfoLevel.Equipment.RackMounts -ge 1) {
                                Section -Style Heading3 -Name 'Rack-Mounts' {

                                    if ($InfoLevel.Equipment.RackMounts -ge 2) {
                                        #region FEX Section
                                        $UcsFex = Get-UcsFex -Ucs $UCSM
                                        if ($UcsFex) {
                                            Section -Style Heading4 -Name 'FEX' {
                                                foreach ($Fex in $UcsFex) {
                                                    Section -Style Heading4 -Name "$($Fex.Rn)" {
                                                        $FabricExtender = [PSCustomObject]@{
                                                            'Name' = $Fex.Rn
                                                            'FEX ID' = $Fex.Id
                                                            'Model' = $Fex.Model
                                                            'Serial' = $Fex.Serial
                                                            'Power' = $Fex.Power
                                                            'Voltage' = $Fex.Voltage
                                                            'Thermal' = $Fex.Thermal
                                                            'License State' = $Fex.LicState
                                                            'State' = $Fex.AdminState 
                                                            'Status' = $Fex.OperState
                                                            'Operability' = $Fex.Operability
                                                        }
                                                        $FabricExtender | Sort-Object 'Name' | Table -List -Name "Fabric Extender $($Fex.Id)" -ColumnWidths 50, 50

                                                        $UcsFexFans = Get-UcsFan -Fex $Fex -Ucs $UCSM
                                                        if ($UcsFexFans) {
                                                            Section -Style Heading4 -Name 'Fans' {
                                                                $FexFans = foreach ($UcsFexFan in $UcsFexFans) {
                                                                    [PSCustomObject]@{
                                                                        'Name' = "Fan $($UcsFexFan.Id)"
                                                                        'Module' = $UcsFexFan.Module
                                                                        'Model' = $UcsFexFan.Model
                                                                        'Serial' = $UcsFexFan.Serial
                                                                        'Power' = $UcsFexFan.Power
                                                                        'Thermal' = $UcsFexFan.Thermal
                                                                        'Presence' = $UcsFexFan.Presence
                                                                        'Operability' = $UcsFexFan.OperState
                                                                    }
                                                                }
                                                                $FexFans | Sort-Object 'Module', 'Name' | Table -Name "Fabric Extender $($Fex.Id) Fans"
                                                            }
                                                        }

                                                        $UcsFexIoms = Get-UcsIom -Fex $Fex -Ucs $UCSM
                                                        if ($UcsFexIoms) {
                                                            Section -Style Heading4 -Name 'IO Modules' {
                                                                $FexIoms = foreach ($UcsFexIom in $UcsFexIoms) {
                                                                    [PSCustomObject]@{
                                                                        'Name' = "IO Module $($UcsFexIom.Id)"
                                                                        'Fabric' = $UcsFexIom.SwitchId
                                                                        'Side' = $UcsFexIom.Side
                                                                        'Model' = $UcsFexIom.Model
                                                                        'Serial' = $UcsFexIom.Serial
                                                                        'Discovery' = $UcsFexIom.Discovery
                                                                        'Config State' = $UcsFexIom.ConfigState 
                                                                        'Operability' = $UcsFexIom.OperState
                                                                        'Thermal' = $UcsFexIom.Thermal
                                                                        'Presence' = $UcsFexIom.Presence
                                                                        'Running Firmware' = (Get-UcsMgmtController -Ucs $UCSM -Dn "$($UcsFexIom.Dn)/mgmt" | Get-UcsFirmwareRunning -Deployment system | Select-Object Version).Version
                                                                        'Backup Firmware' = (Get-UcsMgmtController -Ucs $UCSM -Dn "$($UcsFexIom.Dn)/mgmt" | Get-UcsFirmwareUpdatable | Select-Object Version).Version
                                                                    }
                                                                }
                                                                $FexIoms | Sort-Object 'Name' | Table -Name "Fabric Extender $($Fex.Id) IO Modules"
                                                            }
                                                        }

                                                        $UcsFexPSUs = Get-UcsPsu -Fex $Fex -Ucs $UCSM
                                                        if ($UcsFexPSUs) {
                                                            Section -Style Heading4 -Name 'PSUs' {
                                                                $FexPSUs = foreach ($UcsFexPSU in $UcsFexPSUs) {
                                                                    [PSCustomObject]@{
                                                                        'Name' = "PSU $($UcsFexPSU.Id)"
                                                                        'Model' = $UcsFexPSU.Model
                                                                        'Power' = $UcsFexPSU.Power
                                                                        'Voltage' = $UcsFexPSU.Voltage
                                                                        'Performance' = $UcsFexPSU.Perf
                                                                        'Thermal' = $UcsFexPSU.Thermal
                                                                        'Presence' = $UcsFexPSU.Presence
                                                                        'Operability' = $UcsFexPSU.OperState
                                                                    }
                                                                }
                                                                $FexPSUs | Sort-Object 'Name' | Table -Name "Fabric Extender $($Fex.Id) PSUs"
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        #endregion FEX Section

                                        #region Rack-Mount Servers
                                        if ($UcsRackUnit) {
                                            Section -Style Heading4 -Name 'Servers' {
                                                $UcsRackMountServers = foreach ($RackUnit in $UcsRackUnit) {
                                                    [PSCustomObject]@{
                                                        'Name' = "Server $($RackUnit.Id)"
                                                        'Model' = $RackUnit.Model
                                                        'Serial' = $RackUnit.Serial
                                                        'CPUs' = $RackUnit.NumOfCpus
                                                        'Cores' = $RackUnit.NumOfCores
                                                        'Threads' = $RackUnit.NumOfThreads
                                                        'Memory GB' = $RackUnit.AvailableMemory / 1024 
                                                        'Adapters' = $RackUnit.NumOfAdaptors 
                                                        'NICs' = $RackUnit.NumOfEthHostIfs
                                                        'HBAs' = $RackUnit.NumOfFcHostIfs
                                                        'Power State' = $RackUnit.OperPower
                                                        'Status' = $RackUnit.OperState
                                                        'Assoc State' = $RackUnit.Association
                                                        'Operability' = $RackUnit.Operability 
                                                    } 
                                                }
                                                $UcsRackMountServers | Table -Name 'Rack-Mount Servers'
                                            }
                                        }
                                        #endregion Rack-Mount Servers
                                    }
                                }
                            }
                        }
                        #endregion UCS Rack-Mounts Section
                    
                        <#
                        #region UCS Server Inventory
                        if ($InfoLevel.Equipment.ServerInventory -ge 1) {
                            Section -Style Heading3 -Name 'Server Inventory' {
                                #region UCS Blade Server Inventory
                                if ($UcsBlade) {
                                    Section -Style Heading4 -Name 'Blade Servers' {
                                        $UcsBladeInventory = foreach ($Blade in $UcsBlade) {
                                            [PSCustomObject]@{
                                                'Name' = "Server $($Blade.SlotId)"
                                                'Chassis ID' = $Blade.ChassisID 
                                                'Model' = $Blade.Model
                                                'Serial' = $Blade.Serial
                                                'CPUs' = $Blade.NumOfCpus
                                                'Cores' = $Blade.NumOfCores
                                                'Threads' = $Blade.NumOfThreads 
                                                'Memory GB' = $Blade.AvailableMemory / 1024
                                                'Adapters' = $Blade.NumOfAdaptors 
                                                'NICs' = $Blade.NumOfEthHostIfs
                                                'HBAs' = $Blade.NumOfFcHostIfs
                                                'Power State' = $Blade.OperPower
                                                'Status' = $Blade.OperState
                                                'Assoc State' = $Blade.Association
                                                'Operability' = $Blade.Operability 
                                            } 
                                        }
                                        $UcsBladeInventory | Sort-Object 'Chassis ID', 'Name' | Table -Name 'Server Inventory' 
                                    }
                                }
                                #endregion UCS Blade Server Inventory
                    
                                #region UCS Rack-Mount Server Inventory
                                if ($UcsRackUnit) {
                                    Section -Style Heading4 -Name 'Rack-Mount Servers' {
                                        $UcsRackMountInventory = foreach ($RackUnit in $UcsRackUnit) {
                                            [PSCustomObject]@{
                                                'Name' = "Server $($RackUnit.Id)"
                                                'Model' = $RackUnit.Model
                                                'Serial' = $RackUnit.Serial
                                                'CPUs' = $RackUnit.NumOfCpus
                                                'Cores' = $RackUnit.NumOfCores
                                                'Threads' = $RackUnit.NumOfThreads
                                                'Memory GB' = $RackUnit.AvailableMemory / 1024 
                                                'Adapters' = $RackUnit.NumOfAdaptors 
                                                'NICs' = $RackUnit.NumOfEthHostIfs
                                                'HBAs' = $RackUnit.NumOfFcHostIfs
                                                'Power State' = $RackUnit.OperPower
                                                'Status' = $RackUnit.OperState
                                                'Assoc State' = $RackUnit.Association
                                                'Operability' = $RackUnit.Operability 
                                            } 
                                        }
                                        $UcsRackMountInventory | Table -Name 'Rack-Mount Servers'
                                    }
                                }
                                #endregion UCS Rack-Mount Server Inventory 
                            
                                #region UCS Adaptor Inventory
                                $UcsAdaptorUnit = Get-UcsAdaptorUnit -Ucs $UCSM
                                if ($UcsAdaptorUnit) {
                                    Section -Style Heading3 -Name 'Server Adaptor Inventory' {
                                        $UcsAdaptorInventory = foreach ($AdaptorUnit in $UcsAdaptorUnit) {
                                            [PSCustomObject]@{
                                                'Chassis Id' = $AdaptorUnit.ChassisId 
                                                'Blade Id' = $AdaptorUnit.BladeId 
                                                'Relative Name' = $AdaptorUnit.Rn 
                                                'Model' = $AdaptorUnit.Model
                                            }
                                        }
                                        $UcsAdaptorInventory | Sort-Object 'Chassis Id', 'Blade Id' | Table -Name 'Server Adaptor Inventory' 
                                    }
                                }
                                #endregion UCS Adaptor Inventory
                                
                                $UcsAdaptorUnitExtn = Get-UcsAdaptorUnitExtn -Ucs $UCSM
                                if ($UcsAdaptorUnitExtn) {
                                    Section -Style Heading3 -Name 'Servers with Adaptor Port Expanders' {
                                        $UcsAdaptorUnitExtn = $UcsAdaptorUnitExtn | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Model, Presence
                                        $UcsAdaptorUnitExtn | Table -Name 'Servers with Adaptor Port Expanders' 
                                    }
                                }

                                $UcsProcessorUnit = Get-UcsProcessorUnit -Ucs $UCSM
                                if ($UcsProcessorUnit) {
                                    Section -Style Heading3 -Name 'Server CPU Inventory' {
                                        $UcsProcessorUnit = $UcsProcessorUnit | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, @{L = 'Socket Designation'; E = {$_.SocketDesignation}}, Cores, 
                                        @{L = 'Cores Enabled'; E = {$_.CoresEnabled}}, Threads, Speed, @{L = 'Operability'; E = {$_.OperState}}, Thermal, Model | Where-Object {$_.OperState -ne 'removed'}
                                        $UcsProcessorUnit | Table -Name 'Server CPU Inventory' 
                                    }
                                }

                                $UcsMemoryUnit = Get-UcsMemoryUnit -Ucs $UCSM
                                if ($UcsMemoryUnit) {
                                    Section -Style Heading3 -Name 'Server Memory Inventory' {
                                        $UcsMemoryUnit = $UcsMemoryUnit | Sort-Object Dn, Location | Where-Object {$_.Capacity -ne 'unspecified'} | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Location, Capacity, Clock, @{L = 'Operability'; E = {$_.OperState}}, Model
                                        $UcsMemoryUnit | Table -Name 'Server Memory Inventory' 
                                    }
                                }

                                $UcsStorageController = Get-UcsStorageController -Ucs $UCSM
                                if ($UcsStorageController) {
                                    Section -Style Heading3 -Name 'Server Storage Controller Inventory' {
                                        $UcsStorageController = $UcsStorageController | Sort-Object Dn | Select-Object Vendor, Model
                                        $UcsStorageController | Table -Name 'Server Storage Controller Inventory' 
                                    }
                                }

                                $UcsStorageLocalDisk = Get-UcsStorageLocalDisk -Ucs $UCSM
                                if ($UcsStorageLocalDisk) {
                                    Section -Style Heading3 -Name 'Server Local Disk Inventory' {
                                        $UcsStorageLocalDisk = $UcsStorageLocalDisk | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Model, Size, Serial | Where-Object {$_.Size -ne 'unknown'}
                                        $UcsStorageLocalDisk | Table -Name 'Server Storage Controller Inventory' 
                                    }
                                }
                            }
                        }
                        #endregion UCS Server Inventory
                        #>

                        #region UCS Policies Section
                        if ($InfoLevel.Equipment.Policies -ge 1) {
                            #region Global Policies
                            Section -Style Heading3 -Name 'Global Policies' {
                                #region Chassis/FEX Discovery Policy Section
                                $UcsChassisDiscoveryPolicy = Get-UcsChassisDiscoveryPolicy -Ucs $UCSM
                                if ($UcsChassisDiscoveryPolicy) {
                                    Section -Style Heading4 -Name 'Chassis/FEX Discovery Policy' {
                                        $UcsChassisFexDiscoveryPolicy = [PSCustomObject]@{
                                            'Action' = $UcsChassisDiscoveryPolicy.Action
                                            'Link Aggregation Preference' = $UcsChassisDiscoveryPolicy.LinkAggregationPref
                                            'Backplane Speed Preference' = $UcsChassisDiscoveryPolicy.BackplaneSpeedPref
                                        }
                                        $UcsChassisFexDiscoveryPolicy | Table -List -Name 'Chassis/FEX Discovery Policy' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion Chassis/FEX Discovery Policy Section

                                #region Rack Server Discovery Policy Section
                                $UcsRackServerDiscPolicy = Get-UcsRackServerDiscPolicy -Ucs $UCSM
                                if ($UcsRackServerDiscPolicy) {
                                    Section -Style Heading4 -Name 'Rack Server Discovery Policy' {
                                        $UcsRackServerDiscoveryPolicy = [PSCustomObject]@{
                                            'Action' = $UcsRackServerDiscPolicy.Action
                                            'Scrub Policy' = Switch ($UcsRackServerDiscPolicy.ScrubPolicyName) {
                                                '' { 'not set' }
                                                default { $UcsRackServerDiscPolicy.ScrubPolicyName }
                                            }
                                        }
                                        $UcsRackServerDiscoveryPolicy | Table -List -Name 'Rack Server Discovery Policy' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion Rack Server Discovery Policy Section

                                #region Rack Management Connection Policy Section
                                #TODO
                                #endregion Rack Management Connection Policy Section

                                #region Power Policy Section
                                $UcsPowerControlPolicy = Get-UcsPowerControlPolicy -Ucs $UCSM
                                if ($UcsPowerControlPolicy) {
                                    Section -Style Heading4 -Name 'Power Policy' {
                                        $UcsPowerControlPol = [PSCustomObject]@{
                                            'Redundancy' = $UcsPowerControlPolicy.Redundancy
                                        }
                                        $UcsPowerControlPol | Sort-Object 'Chassis' | Table -List -Name 'Power Policy' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion Power Policy Section

                                #region Global Power Allocation Policy Section
                                $UcsPowerMgmtPolicy = Get-UcsPowerMgmtPolicy -Ucs $UCSM
                                if ($UcsPowerMgmtPolicy) {
                                    Section -Style Heading4 -Name 'Global Power Allocation Policy' {
                                        $UcsGlobalPowerPolicy = [PSCustomObject]@{
                                            'Allocation Method' = Switch ($UcsPowerMgmtPolicy.Style) {
                                                'intelligent-policy-driven' { 'Policy Driven Chassis Group Cap' }
                                                'manual-per-blade' { 'Manual Blade Level Cap' }
                                            }
                                        }
                                        $UcsGlobalPowerPolicy | Table -List -Name 'Global Power Allocation Policy' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion Global Power Allocation Policy Section
                            }
                            #endregion Global Policies Section
                        }
                        #endregion UCS Policies Section
            
                        <#
                        Section -Style Heading3 -Name 'Firmware' {
                            $UcsFirmware = Get-UcsFirmwareRunning
                            if ($UcsFirmware) {
                                Section -Style Heading2 -Name 'Firmware Management' {
                                    Section -Style Heading3 -Name 'UCS Manager' {
                                        $UcsmFirmware = $UcsFirmware | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Type, Version | Sort-Object Dn | Where-Object {$_.Type -eq 'mgmt-ext'}
                                        $UcsmFirmware | Table -Name 'UCS Manager Firmware' 
                                    }

                                    Section -Style Heading3 -Name 'Fabric Interconnect' {
                                        $UcsFiFirmware = $UcsFirmware | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Type, Version | Sort-Object Dn | Where-Object {$_.Type -eq 'switch-kernel' -OR $_.Type -eq 'switch-software'}
                                        $UcsFiFirmware | Table -Name 'Fabric Interconnect Firmware' 
                                    }

                                    Section -Style Heading3 -Name 'IOM' {
                                        $UcsIomFiFirmware = $UcsFirmware | Sort-Object Dn | Select-Object Deployment, @{L = 'Distinguished Name'; E = {$_.Dn}}, Type, Version | Where-Object {$_.Type -eq 'iocard'} | Where-Object -FilterScript {$_.Deployment -notlike 'boot-loader'}
                                        $UcsIomFiFirmware | Table -Name 'IOM Firmware' 
                                    }

                                    Section -Style Heading3 -Name 'Server Adapters' {
                                        $UcsServerAdapterFirmware = $UcsFirmware | Sort-Object Dn | Select-Object Deployment, @{L = 'Distinguished Name'; E = {$_.Dn}}, Type, Version | Where-Object {$_.Type -eq 'adaptor'} | Where-Object -FilterScript {$_.Deployment -notlike 'boot-loader'}
                                        $UcsServerAdapterFirmware | Table -Name 'Server Adapter Firmware' 
                                    }

                                    Section -Style Heading3 -Name 'Server CIMC' {
                                        $UcsServerCimcFirmware = $UcsFirmware | Sort-Object Dn | Select-Object Deployment, @{L = 'Distinguished Name'; E = {$_.Dn}}, Type, Version | Where-Object {$_.Type -eq 'blade-controller'} | Where-Object -FilterScript {$_.Deployment -notlike 'boot-loader'}
                                        $UcsServerCimcFirmware | Table -Name 'Server CIMC Firmware' 
                                    }

                                    Section -Style Heading3 -Name 'Server BIOS' {
                                        $UcsServerBios = $UcsFirmware | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Type, Version | Where-Object {$_.Type -eq 'blade-bios'}
                                        $UcsServerBios | Table -Name 'Server BIOS' 
                                    }

                                    Section -Style Heading3 -Name 'Host Firmware Packages' {
                                        $UcsFirmwareComputeHostPack = Get-UcsFirmwareComputeHostPack | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Name, @{L = 'Blade Bundle Version'; E = {$_.BladeBundleVersion}}, @{L = 'Rack Bundle Version'; E = {$_.RackBundleVersion}}
                                        $UcsFirmwareComputeHostPack | Table -Name 'Host Firmware Packages' 
                                    }
                                }
                            }
                        }
                        #>
                    }
                }
                #endregion Equipment Section

                #region Servers Section
                if ($Section.Servers) {
                    Section -Style Heading2 -Name 'Servers' {
                        #region Service Profiles
                        $UcsServiceProfiles = Get-UcsServiceProfile | Where-Object { $_.Type -eq 'instance' } | Sort-Object Name
                        if ($UcsServiceProfiles) {
                            #region Service Profiles Section
                            Section -Style Heading3 -Name 'Service Profiles' {
                                foreach ($ServiceProfile in $UcsServiceProfiles) {
                                    Section -Style Heading4 -Name "$($ServiceProfile.Name)" {
                                        $ServiceProfiles = [PSCustomObject]@{
                                            'Name' = $ServiceProfile.Name
                                            'User Label' = $ServiceProfile.UsrLbl
                                            'Description' = $ServiceProfile.Descr
                                            'Distinguished Name' = $ServiceProfile.Dn
                                            'Unique Identifier' = $ServiceProfile.Uuid
                                            'UUID Pool' = $ServiceProfile.IdentPoolName
                                            'Associated Server' = $ServiceProfile.PnDn
                                            'Service Profile Template' = $ServiceProfile.SrcTemplName
                                            'Associated State' = $ServiceProfile.AssocState
                                            'Assigned State' = $ServiceProfile.ConfigState
                                            'Status' = $ServiceProfile.OperState
                                            'Boot Policy Name' = Switch ($ServiceProfile.BootPolicyName) {
                                                '' { 'not set' }
                                                default { $ServiceProfile.BootPolicyName }
                                            }
                                            'Host Firmware Policy Name' = Switch ($ServiceProfile.HostFwPolicyName) {
                                                '' { 'not set' }
                                                default { $ServiceProfile.HostFwPolicyName }
                                            }
                                            'IPMI Access Profile Policy' = Switch ($ServiceProfile.MgmtAccessPolicyName) {
                                                '' { 'not set' }
                                                default { $ServiceProfile.MgmtAccessPolicyName }
                                            }
                                            'Local Disk Policy Name' = Switch ($ServiceProfile.LocalDiskPolicyName) {
                                                '' { 'not set' }
                                                default { $ServiceProfile.LocalDiskPolicyName }
                                            }
                                            'Maintenance Policy Name' = Switch ($ServiceProfile.MaintPolicyName) {
                                                '' { 'not set' }
                                                default { $ServiceProfile.MaintPolicyName }
                                            }
                                            'Power Control Policy' = Switch ($ServiceProfile.PowerPolicyName) {
                                                '' { 'not set' }
                                                default { $ServiceProfile.PowerPolicyName }
                                            }
                                            'Scrub Policy' = Switch ($ServiceProfile.ScrubPolicyName) {
                                                '' { 'not set' }
                                                default { $ServiceProfile.ScrubPolicyName }
                                            }
                                            'Stats Policy' = Switch ($ServiceProfile.StatsPolicyName) {
                                                '' { 'not set' }
                                                default { $ServiceProfile.StatsPolicyName }
                                            }
                                            'KVM Management Policy' = Switch ($ServiceProfile.KvmMgmtPolicyName) {
                                                '' { 'not set' }
                                                default { $ServiceProfile.KvmMgmtPolicyName }
                                            }
                                            'Power Sync Policy' = Switch ($ServiceProfile.PowerSyncPolicyName) {
                                                '' { 'not set' }
                                                default { $ServiceProfile.PowerSyncPolicyName }
                                            }
                                            'Graphics Card Policy' = Switch ($ServiceProfile.GraphicsCardPolicyName) {
                                                '' { 'not set' }
                                                default { $ServiceProfile.GraphicsCardPolicyName }
                                            }
                                        }
                                        $ServiceProfiles | Table -List -Name "$($ServiceProfile.Name) Service Profile" -ColumnWidths 50, 50
                                    }
                                }
                            }
                            #endregion Service Profiles Section
                        }
                        #endregion Service Profiles

                        #region Service Profile Templates
                        $UcsServiceProfileTemplates = Get-UcsServiceProfile | Where-Object { $_.Type -ne 'instance' } | Sort-Object Name
                        if ($UcsServiceProfileTemplates) {
                            #region Service Profile Templates Section
                            Section -Style Heading3 -Name 'Service Profile Templates' {
                                foreach ($ServiceProfileTemplate in $UcsServiceProfileTemplates) {
                                    Section -Style Heading4 -Name "$($ServiceProfileTemplate.Name)" {
                                        $ServiceProfileTemplates = [PSCustomObject]@{
                                            'Name' = $ServiceProfileTemplate.Name
                                            'User Label' = $ServiceProfileTemplate.UsrLbl
                                            'Description' = $ServiceProfileTemplate.Descr
                                            'Distinguished Name' = $ServiceProfileTemplate.Dn
                                            'Unique Identifier' = $ServiceProfileTemplate.Uuid
                                            'UUID Pool' = $ServiceProfileTemplate.IdentPoolName
                                            'Associated Server' = $ServiceProfileTemplate.PnDn
                                            'Service Profile Template' = $ServiceProfileTemplate.SrcTemplName
                                            'Associated State' = $ServiceProfileTemplate.AssocState
                                            'Assigned State' = $ServiceProfileTemplate.ConfigState
                                            'Status' = $ServiceProfileTemplate.OperState
                                            'Boot Policy Name' = $ServiceProfileTemplate.BootPolicyName 
                                            'Host Firmware Policy Name' = $ServiceProfileTemplate.HostFwPolicyName
                                            'IPMI Access Profile Policy' = $ServiceProfileTemplate.MgmtAccessPolicyName
                                            'Local Disk Policy Name' = $ServiceProfileTemplate.LocalDiskPolicyName 
                                            'Maintenance Policy Name' = $ServiceProfileTemplate.MaintPolicyName 
                                            'Power Control Policy' = $ServiceProfileTemplate.PowerPolicyName
                                            'Scrub Policy' = $ServiceProfileTemplate.ScrubPolicyName
                                            'Stats Policy' = $ServiceProfileTemplate.StatsPolicyName
                                            'KVM Management Policy' = $ServiceProfileTemplate.KvmMgmtPolicyName
                                            'Power Sync Policy' = $ServiceProfileTemplate.PowerSyncPolicyName
                                            'Graphics Card Policy' = $ServiceProfileTemplate.GraphicsCardPolicyName
                                        }
                                        $ServiceProfileTemplates | Table -List -Name "$($ServiceProfileTemplate.Name) Service Profile Template" -ColumnWidths 50, 50
                                    }
                                }
                            }
                            #endregion Service Profile Templates Section
                        }
                        #endregion Service Profile Templates

                        #region Server Policies Section
                        Section -Style Heading2 -Name 'Policies' {

                            Section -Style Heading3 -Name 'Adapter Policies' {
                                #region Ethernet Adapter Policy
                                $UcsEthAdapterPolicy = Get-UcsEthAdapterPolicy -Ucs $UCSM
                                if ($UcsEthAdapterPolicy) {
                                    Section -Style Heading4 -Name 'Ethernet Adapter Policies' {
                                        $EthAdapterPolicies = foreach ($EthAdapterPolicy in $UcsEthAdapterPolicy) {
                                            [PSCustomObject]@{
                                                'Name' = $EthAdapterPolicy.Name
                                                'Distinguished Name' = $EthAdapterPolicy.Dn
                                                'Description' = $EthAdapterPolicy.Descr
                                                'Owner' = $EthAdapterPolicy.PolicyOwner
                                            }
                                        }
                                        $EthAdapterPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Ethernet Adapter Policies'
                                    }
                                }
                                #endregion Ethernet Adapter Policy

                                #region iSCSI Adapter Policy
                                $UcsiScsiAdapterPolicy = Get-UcsIScsiAdapterPolicy -Ucs $UCSM
                                if ($UcsiScsiAdapterPolicy) {
                                    Section -Style Heading4 -Name 'iSCSI Adapter Policies' {
                                        $iScsiAdapterPolicies = foreach ($iScsiAdapterPolicy in $UcsiScsiAdapterPolicy ) {
                                            [PSCustomObject]@{
                                                'Name' = $iScsiAdapterPolicy.Name
                                                'Distinguished Name' = $iScsiAdapterPolicy.Dn
                                                'Description' = $iScsiAdapterPolicy.Descr
                                                'Owner' = $iScsiAdapterPolicy.PolicyOwner
                                            }
                                        }
                                        $iScsiAdapterPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'iSCSI Adapter Policies'
                                    }
                                }
                                #endregion iSCSI Adapter Policy

                                #region Fibre Channel Adapter Policy
                                $UcsFcAdapterPolicy = Get-UcsFcAdapterPolicy -Ucs $UCSM
                                if ($UcsFcAdapterPolicy) {
                                    Section -Style Heading4 -Name 'Fibre Channel Adapter Policies' {
                                        $FcAdapterPolicies = foreach ($FcAdapterPolicy in $UcsFcAdapterPolicy) {
                                            [PSCustomObject]@{
                                                'Name' = $FcAdapterPolicy.Name
                                                'Distinguished Name' = $FcAdapterPolicy.Dn
                                                'Description' = $FcAdapterPolicy.Descr
                                                'Owner' = $FcAdapterPolicy.PolicyOwner
                                            }
                                        }
                                        $FcAdapterPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Fibre Channel Adapter Policies'
                                    }
                                }
                                #endregion Fibre Channel Adapter Policy
                            }

                            #region BIOS Policies
                            $UcsBiosPolicy = Get-UcsBiosPolicy -Ucs $UCSM
                            if ($UcsBiosPolicy) {
                                Section -Style Heading3 -Name 'BIOS Policies' {
                                    $UcsBiosPolicies = foreach ($BiosPolicy in $UcsBiosPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $BiosPolicy.Name
                                            'Distinguished Name' = $BiosPolicy.Dn
                                            'Description' = $BiosPolicy.Descr
                                            'Owner' = $BiosPolicy.PolicyOwner
                                            'Reboot on BIOS Settings Change' = $BiosPolicy.RebootOnUpdate
                                        }
                                    }
                                    $UcsBiosPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'BIOS Policies'
                                }
                            }

                            <#                        
                            Section -Style Heading4 -Name 'BIOS Policy Settings' {
                                Get-UcsBiosPolicy | Get-UcsBiosVfQuietBoot | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfQuietBoot'
                                Get-UcsBiosPolicy | Get-UcsBiosVfPOSTErrorPause | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfPOSTErrorPause'
                                Get-UcsBiosPolicy | Get-UcsBiosVfResumeOnACPowerLoss | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfResumeOnACPowerLoss'
                                Get-UcsBiosPolicy | Get-UcsBiosVfFrontPanelLockout | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfFrontPanelLockout'
                                Get-UcsBiosPolicy | Get-UcsBiosTurboBoost | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy TurboBoost'
                                Get-UcsBiosPolicy | Get-UcsBiosEnhancedIntelSpeedStep | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy EnhancedIntelSpeedStep'
                                Get-UcsBiosPolicy | Get-UcsBiosHyperThreading | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy HyperThreading'
                                Get-UcsBiosPolicy | Get-UcsBiosVfCoreMultiProcessing | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfCoreMultiProcessing'
                                Get-UcsBiosPolicy | Get-UcsBiosExecuteDisabledBit | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy ExecuteDisabledBit'
                                Get-UcsBiosPolicy | Get-UcsBiosVfIntelVirtualizationTechnology | Sort-Object  Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfIntelVirtualizationTechnology'
                                Get-UcsBiosPolicy | Get-UcsBiosVfDirectCacheAccess | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfDirectCacheAccess'
                                Get-UcsBiosPolicy | Get-UcsBiosVfProcessorCState | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfProcessorCState'
                                Get-UcsBiosPolicy | Get-UcsBiosVfProcessorC1E | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfProcessorC1E'
                                Get-UcsBiosPolicy | Get-UcsBiosVfProcessorC3Report | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfProcessorC3Report'
                                Get-UcsBiosPolicy | Get-UcsBiosVfProcessorC6Report | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfProcessorC6Report'
                                Get-UcsBiosPolicy | Get-UcsBiosVfProcessorC7Report | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfProcessorC7Report'
                                Get-UcsBiosPolicy | Get-UcsBiosVfCPUPerformance | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfCPUPerformance'
                                Get-UcsBiosPolicy | Get-UcsBiosVfMaxVariableMTRRSetting | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfMaxVariableMTRRSetting'
                                Get-UcsBiosPolicy | Get-UcsBiosIntelDirectedIO | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy IntelDirectedIO'
                                Get-UcsBiosPolicy | Get-UcsBiosVfSelectMemoryRASConfiguration | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfSelectMemoryRASConfiguration'
                                Get-UcsBiosPolicy | Get-UcsBiosNUMA | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy NUMA'
                                Get-UcsBiosPolicy | Get-UcsBiosLvDdrMode | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy LvDdrMode'
                                Get-UcsBiosPolicy | Get-UcsBiosVfUSBBootConfig | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfUSBBootConfig'
                                Get-UcsBiosPolicy | Get-UcsBiosVfUSBFrontPanelAccessLock | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfUSBFrontPanelAccessLock'
                                Get-UcsBiosPolicy | Get-UcsBiosVfUSBSystemIdlePowerOptimizingSetting | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfUSBSystemIdlePowerOptimizingSetting'
                                Get-UcsBiosPolicy | Get-UcsBiosVfMaximumMemoryBelow4GB | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfMaximumMemoryBelow4GB'
                                Get-UcsBiosPolicy | Get-UcsBiosVfMemoryMappedIOAbove4GB | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfMemoryMappedIOAbove4GB'
                                Get-UcsBiosPolicy | Get-UcsBiosVfBootOptionRetry | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfBootOptionRetry'
                                Get-UcsBiosPolicy | Get-UcsBiosVfIntelEntrySASRAIDModule | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfIntelEntrySASRAIDModule'
                                Get-UcsBiosPolicy | Get-UcsBiosVfOSBootWatchdogTimer | Sort-Object Dn | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Vp* | Table -Name 'BIOS Policy VfOSBootWatchdogTimer'
                            }
                            #>
                            #endregion BIOS Policies

                            #region Boot Policies
                            $UcsBootPolicy = Get-UcsBootPolicy -Ucs $UCSM
                            if ($UcsBootPolicy) {
                                Section -Style Heading3 -Name 'Boot Policies' {
                                    $UcsBootPolicies = foreach ($BootPolicy in $UcsBootPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $BootPolicy.Name
                                            'Distinguished Name' = $BootPolicy.Dn
                                            'Description' = $BootPolicy.Descr
                                            'Owner' = $BootPolicy.PolicyOwner
                                            'Reboot on Boot Order Change' = $BootPolicy.RebootOnUpdate
                                            'Enforce vNIC/vHBA/iSCSI Name' = $BootPolicy.EnforceVnicName
                                            'Boot Mode' = $BootPolicy.BootMode
                                        }
                                    }
                                    $UcsBootPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Boot Policies' 
                                }
                            }
                            #endregion Boot Policies

                            #region Diagnostic Policies
                            $UcsDiagRunPolicy = Get-UcsDiagRunPolicy -Ucs $UCSM
                            if ($UcsDiagRunPolicy) {
                                Section -Style Heading3 -Name 'Diagnostic Policies' {
                                    $UcsDiagRunPolicies = foreach ($DiagRunPolicy in $UcsDiagRunPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $DiagRunPolicy.Name
                                            'Distinguished Name' = $DiagRunPolicy.Dn
                                            'Description' = $DiagRunPolicy.Descr
                                            'Owner' = $DiagRunPolicy.PolicyOwner
                                        }
                                    }
                                    $UcsDiagRunPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Diagnostic Policies'
                                }
                            }
                            #endregion Diagnostic Policies

                            #region Graphics Card Policies
                            $UcsComputeGraphicsCardPolicy = Get-UcsComputeGraphicsCardPolicy -Ucs $UCSM
                            if ($UcsComputeGraphicsCardPolicy) {
                                Section -Style Heading3 -Name 'Graphics Card Policies' {
                                    $UcsGraphicsCardPolicies = foreach ($ComputeGraphicsCardPolicy in $UcsComputeGraphicsCardPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $ComputeGraphicsCardPolicy.Name
                                            'Distinguished Name' = $ComputeGraphicsCardPolicy.Dn
                                            'Description' = $ComputeGraphicsCardPolicy.Descr
                                            'Owner' = $ComputeGraphicsCardPolicy.PolicyOwner
                                            'Graphics Card Mode' = $ComputeGraphicsCardPolicy.GraphicsCardMode
                                        }
                                    }
                                    $UcsGraphicsCardPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Graphics Card Policies'
                                }
                            }
                            #endregion Graphics Card Policies

                            <#
                            #region Host Firmware Packages
                            Section -Style Heading3 -Name 'Host Firmware Packages' {
                            }
                            #endregion Host Firmware Packages

                            #region IPMI Access Profiles
                            Section -Style Heading3 -Name 'IPMI Access Profiles' {
                            }
                            #endregion IPMI Access Profiles
                            #>

                            #region KVM Management Policies
                            $UcsComputeKvmMgmtPolicy = Get-UcsComputeKvmMgmtPolicy -Ucs $UCSM | Where-Object { $_.Name -ne 'policy' }
                            if ($UcsComputeKvmMgmtPolicy) {
                                Section -Style Heading3 -Name 'KVM Management Policies' {
                                    $ComputeKvmMgmtPolicies = foreach ($ComputeKvmMgmtPolicy in $UcsComputeKvmMgmtPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $ComputeKvmMgmtPolicy.Name
                                            'Distinguished Name' = $ComputeKvmMgmtPolicy.Dn
                                            'Description' = $ComputeKvmMgmtPolicy.Descr
                                            'Owner' = $ComputeKvmMgmtPolicy.PolicyOwner
                                            'vMedia Encryption' = $ComputeKvmMgmtPolicy.VmediaEncryption
                                        }
                                    }
                                    $ComputeKvmMgmtPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'KVM Management Policies'
                                }
                            }
                            #endregion KVM Management Policies

                            #region Local Disk Config Policies
                            $UcsLocalDiskConfigPolicy = Get-UcsLocalDiskConfigPolicy -Ucs $UCSM
                            if ($UcsLocalDiskConfigPolicy) {
                                Section -Style Heading3 -Name 'Local Disk Config Policies' {
                                    $LocalDiskConfigPolicies = foreach ($LocalDiskConfigPolicy in $UcsLocalDiskConfigPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $LocalDiskConfigPolicy.Name
                                            'Distinguished Name' = $LocalDiskConfigPolicy.Dn
                                            'Description' = $LocalDiskConfigPolicy.Descr
                                            'Owner' = $LocalDiskConfigPolicy.PolicyOwner
                                            'Mode' = $LocalDiskConfigPolicy.Mode
                                            'Protect Configuration' = $LocalDiskConfigPolicy.ProtectConfig
                                            'FlexFlash State' = $LocalDiskConfigPolicy.FlexFlashState
                                            'FlexFlash RAID Reporting State' = $LocalDiskConfigPolicy.FlexFlashRAIDReportingState
                                            'FlexFlash Removable State' = $LocalDiskConfigPolicy.FlexFlashRemovableState
                                        }
                                    }
                                    $LocalDiskConfigPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Local Disk Config Policies' 
                                }
                            }
                            #endregion Local Disk Config Policies

                            #region Maintenance Policies
                            $UcsMaintenancePolicy = Get-UcsMaintenancePolicy -Ucs $UCSM
                            if ($UcsMaintenancePolicy) {
                                Section -Style Heading3 -Name 'Maintenance Policies' {
                                    $MaintenancePolicies = foreach ($MaintenancePolicy in $UcsMaintenancePolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $MaintenancePolicy.Name
                                            'Distinguished Name' = $MaintenancePolicy.Dn
                                            'Description' = $MaintenancePolicy.Descr
                                            'Owner' = $MaintenancePolicy.PolicyOwner
                                            'Soft Shutdown Timer' = $MaintenancePolicy.SoftShutdownTimer
                                            'Storage Config Development Policy' = $MaintenancePolicy.DataDisr
                                            'Reboot Policy' = $MaintenancePolicy.UptimeDisr
                                        }
                                    }
                                    $MaintenancePolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Maintenance Policies' 
                                }
                            }
                            #endregion Maintenance Policies
            
                            <#
                            #region Management Firmware Packages
                            Section -Style Heading3 -Name 'Management Firmware Packages' {
                            }
                            #endregion Management Firmware Packages
                            #>

                            #region Memory Policy
                            $UcsComputeMemoryConfigPolicy = Get-UcsComputeMemoryConfigPolicy -Ucs $UCSM
                            if ($UcsComputeMemoryConfigPolicy) {
                                Section -Style Heading3 -Name 'Memory Policy' {
                                    $ComputeMemoryConfigPolicies = foreach ($ComputeMemoryConfigPolicy in $UcsComputeMemoryConfigPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $ComputeMemoryConfigPolicy.Name
                                            'Distinguished Name' = $ComputeMemoryConfigPolicy.Dn
                                            'Description' = $ComputeMemoryConfigPolicy.Descr
                                            'Owner' = $ComputeMemoryConfigPolicy.PolicyOwner
                                            'Blacklisting' = $ComputeMemoryConfigPolicy.Blacklisting
                                        }
                                    }
                                    $ComputeMemoryConfigPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Memory Policy'
                                }
                            }
                            #endregion Memory Policy
                            
                            #region Power Control Policies
                            $UcsPowerPolicy = Get-UcsPowerPolicy -Ucs $UCSM
                            if ($UcsPowerPolicy) {
                                Section -Style Heading3 -Name 'Power Control Policies' {
                                    $PowerPolicies = foreach ($PowerPolicy in $UcsPowerPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $PowerPolicy.Name
                                            'Distinguised Name' = $PowerPolicy.Dn
                                            'Description' = $PowerPolicy.Descr
                                            'Owner' = $PowerPolicy.PolicyOwner
                                            'Fan Speed Policy' = $PowerPolicy.FanSpeed
                                        }
                                    }
                                    $PowerPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Power Control Policies'
                                }
                            }
                            #endregion Power Control Policies
                            
                            #region Power Sync Policies
                            $UcsComputePowerSyncPolicy = Get-UcsComputePowerSyncPolicy -Ucs $UCSM
                            if ($UcsComputePowerSyncPolicy) {
                                Section -Style Heading3 -Name 'Power Sync Policies' {
                                    $ComputePowerSyncPolicies = foreach ($ComputePowerSyncPolicy in $UcsComputePowerSyncPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $ComputePowerSyncPolicy.Name
                                            'Distinguished Name' = $ComputePowerSyncPolicy.Dn
                                            'Description' = $ComputePowerSyncPolicy.Descr
                                            'Owner' = $ComputePowerSyncPolicy.PolicyOwner
                                            'Sync Option' = $ComputePowerSyncPolicy.SyncOption
                                        }
                                    }
                                    $ComputePowerSyncPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Power Sync Policies'
                                }
                            }
                            #endregion Power Sync Policies

                            #region Scrub Policies
                            $UcsScrubPolicy = Get-UcsScrubPolicy -Ucs $UCSM | Where-Object { $_.Name -ne 'policy' }
                            if ($UcsScrubPolicy) {
                                Section -Style Heading3 -Name 'Scrub Policies' {
                                    $UcsScrubPolicies = foreach ($ScrubPolicy in $UcsScrubPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $ScrubPolicy.Name
                                            'Distinguished Name' = $ScrubPolicy.Dn
                                            'Description' = $ScrubPolicy.Descr
                                            'Owner' = $ScrubPolicy.PolicyOwner
                                            'Disk Scrub' = $ScrubPolicy.DiskScrub
                                            'BIOS Settings Scrub' = $ScrubPolicy.BiosSettingsScrub
                                            'FlexFlash Scrub' = $ScrubPolicy.FlexFlashScrub
                                        }
                                    }
                                    $UcsScrubPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Scrub Policies' 
                                }
                            }
                            #endregion Scrub Policies

                            <#
                            #region Serial over LAN Policies
                            Section -Style Heading3 -Name 'Serial over LAN Policies' {
                            }
                            #endregion Serial over LAN Policies

                            #region Server Pool Policies
                            Section -Style Heading3 -Name 'Server Pool Policies' {
                            }
                            #endregion Server Pool Policies
                            #>
                            
                            #region Server Pool Policy Qualifications
                            $UcsServerPoolQualification = Get-UcsServerPoolQualification -Ucs $UCSM
                            if ($UcsServerPoolQualification) {
                                Section -Style Heading3 -Name 'Server Pool Policy Qualifications' {
                                    $ServerPoolQualifications = foreach ($ServerPoolQualification in $UcsServerPoolQualification) {
                                        [PSCustomObject]@{
                                            'Name' = $ServerPoolQualification.Name
                                            'Distinguished Name' = $ServerPoolQualification.Dn
                                            'Description' = $ServerPoolQualification.Descr
                                            'Owner' = $ServerPoolQualification.PolicyOwner
                                        }
                                    }
                                    $ServerPoolQualifications | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Server Pool Policy Qualifications'
                                }
                            }
                            #endregion Server Pool Policy Qualifications

                            <#
                            #region Threshold Policies
                            Section -Style Heading3 -Name 'Threshold Policies' {
                            }
                            #endregion Threshold Policies

                            #region iSCSI Authentication Profiles
                            Section -Style Heading3 -Name 'iSCSI Authentication Profiles' {
                            }
                            #endregion iSCSI Authentication Profiles

                            #region vMedia Policies
                            Section -Style Heading3 -Name 'vMedia Policies' {
                            }
                            #endregion vMedia Policies
                            #>

                        }
                        #endregion Server Policies Section


                        #region Server Pools Section
                        Section -Style Heading2 -Name 'Pools' {
                            #region UUID Pools
                            $UcsUuidSuffixPool = Get-UcsUuidSuffixPool -Ucs $UCSM | Where-Object { $_.Size -gt 0 }
                            if ($UcsUuidSuffixPool) {
                                Section -Style Heading3 -Name 'UUID Suffix Pools' {
                                    $UuidSuffixPool = foreach ($UuidPool in $UcsUuidSuffixPool) {
                                        $UcsUuidSuffixBlock = Get-UcsUuidSuffixBlock -UuidSuffixPool $UuidPool 
                                        [PSCustomObject]@{
                                            'Name' = $UuidPool.Name
                                            'Owner' = $UuidPool.PolicyOwner
                                            'Description' = $UuidPool.Descr
                                            'Size' = $UuidPool.Size
                                            'Assigned' = $UuidPool.Assigned
                                            'Assignment Order' = $UuidPool.AssignmentOrder
                                            'UUID Suffix Blocks' = ($UcsUuidSuffixBlock | Sort-Object From | ForEach-Object { "$($_.From) - $($_.To)" }) -join [Environment]::NewLine
                                        }
                                    }
                                    $UuidSuffixPool | Table -Name 'UUID Suffix Pools' 
                                }
                            }
                            #endregion UUID Pools

                            #region UUID Suffixes
                            $UcsUuidPoolAddr = Get-UcsUuidpoolAddr -Ucs $UCSM | Where-Object { $_.Assigned -eq 'yes' }
                            if ($UcsUuidPoolAddr) {
                                Section -Style Heading3 -Name 'UUID Suffixes' {
                                    $UuidPoolAddr = foreach ($UuidAddr in $UcsUuidPoolAddr) {
                                        [PSCustomObject]@{
                                            'Name' = $UuidAddr.Id
                                            'Owner' = $UuidAddr.Owner
                                            'Assigned' = $UuidAddr.Assigned
                                            'Assigned To' = $UuidAddr.AssignedToDn
                                        }
                                    }
                                    $UuidPoolAddr | Table -Name 'UUID Suffixes' 
                                }
                            }
                            #endregion UUID Suffixes

                            #region Server Pools
                            $UcsServerPool = Get-UcsServerPool -Ucs $UCSM | Where-Object { $_.Size -gt 0 }
                            if ($UcsServerPool) {
                                Section -Style Heading3 -Name 'Server Pools' {
                                    $ServerPool = foreach ($Server in $UcsServerPool) {
                                        [PSCustomObject]@{
                                            'Name' = $Server.Name
                                            'Owner' = $Server.PolicyOwner
                                            'Size' = $Server.Size
                                            'Assigned' = $Server.Assigned
                                            'Assignment Order' = $Server.AssignmentOrder
                                        }
                                    }
                                    $ServerPool | Table -Name 'Server Pools' 
                                }
                            }
                            #endregion Server Pools

                            #region Server Pool Assignments
                            Section -Style Heading3 -Name 'Server Pool Assignments' {
                                #    $UcsComputePooledSlot = Get-UcsComputePooledSlot | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Rn
                                #    $UcsComputePooledSlot | Table -Name 'Server Pool Assignments' 
                            }
                            #endregion Server Pool Assignments
                        }
                        #endregion Server Pools Section
                    }
                }
                #endregion Servers Section

                #region LAN Section
                if ($Section.LAN) {
                    Section -Style Heading2 -Name 'LAN' {
                        #region LAN Cloud Section
                        Section -Style Heading3 -Name 'LAN Cloud' {

                            $UcsLanCloud = Get-UcsLanCloud -Ucs $UCSM
                            if ($UcsLanCloud) {
                                $LanCloud = [PSCustomObject]@{
                                    'UCS' = $UcsLanCloud.Ucs
                                    'Mode' = $UcsLanCloud.Mode
                                } 
                                $LanCloud | Table -Name 'LAN Cloud' -ColumnWidths 50, 50
                            }

                            #region Port Channels and Uplinks Section
                            $UcsUplinkPortChannel = Get-UcsUplinkPortChannel -Ucs $UCSM
                            if ($UcsUplinkPortChannel) {
                                Section -Style Heading4 -Name 'Port Channels and Uplinks' {
                                    $UplinkPortChannel = foreach ($PortChannel in $UcsUplinkPortChannel) {
                                        [PSCustomObject]@{
                                            'Fabric' = $PortChannel.SwitchId
                                            'ID' = $PortChannel.PortId
                                            'Name' = $PortChannel.Name
                                            'Description' = $PortChannel.Descr
                                            'State' = $PortChannel.AdminState
                                            'If Type' = $PortChannel.IfType
                                            'Transport' = $PortChannel.Transport
                                            'Flow Control Policy' = $PortChannel.FlowCtrlPolicy
                                            'LACP Policy' = $PortChannel.LacpPolicyName
                                            'Speed' = $PortChannel.AdminSpeed
                                            'Operational Speed' = "$($PortChannel.Bandwidth)gbps"
                                            'Status' = $PortChannel.OperState
                                        }
                                    }
                                    $UplinkPortChannel | Sort-Object 'Fabric', 'ID' | Table -Name 'Port Channels and Uplinks'
                                    BlankLine
                                    foreach ($PortChannel in $UcsUplinkPortChannel) {
                                        $UcsUplinkPortChannelMember = Get-UcsUplinkPortChannelMember -Ucs $UCSM -UplinkPortChannel $PortChannel
                                        $UplinkPortChannelMember = foreach ($PortChannelMember in $UcsUplinkPortChannelMember) {
                                            [PSCustomObject]@{
                                                'Fabric' = $PortChannelMember.SwitchId
                                                'Port' = "$($PortChannelMember.SlotId)/$($PortChannelMember.PortId)"
                                                'Port Channel' = "$($PortChannel.PortId) $($PortChannel.Name)"
                                                'State' = $PortChannelMember.AdminState
                                                'Membership' = $PortChannelMember.Membership
                                                'Transport' = $PortChannelMember.Transport
                                                'Medium' = $PortChannelMember.Type
                                                'If Role' = $PortChannelMember.IfRole
                                                'If Type' = $PortChannelMember.IfType
                                                'Locale' = $PortChannelMember.Locale
                                                'Link Profile' = $PortChannelMember.EthLinkProfileName
                                                'Status' = $PortChannelMember.OperState
                                            }
                                        }
                                        $UplinkPortChannelMembers += $UplinkPortChannelMember
                                    }
                                    $UplinkPortChannelMembers | Sort-Object 'Fabric', 'Port', 'Port Channel' | Table -Name 'Port Channel Members'

                                }
                            }
                            #endregion Port Channels and Uplinks Section

                            #region VLANs
                            $UcsVlan = Get-UcsVlan | Where-Object { $_.IfRole -eq 'Network' }
                            if ($UcsVlan) {
                                Section -Style Heading4 -Name 'VLANs' {
                                    $Vlans = foreach ($Vlan in $UcsVlan) {
                                        [PSCustomObject]@{
                                            'Fabric' = $Vlan.SwitchId
                                            'VLAN Name' = $Vlan.Name
                                            'VLAN ID' = $Vlan.Id
                                            'Native' = $Vlan.DefaultNet
                                            'Type' = $Vlan.Type
                                            'Transport' = $Vlan.Transport
                                            'Locale' = $Vlan.Locale
                                            'Sharing Type' = $Vlan.Sharing
                                            'Multicast Policy Name' = Switch ($Vlan.McastPolicyName) {
                                                '' { 'not set' }
                                                default { $Vlan.McastPolicyName }
                                            }
                                            'Multicast Policy Instance' = $Vlan.OperMcastPolicyName
                                        }
                                    }
                                    $Vlans | Sort-Object 'Fabric', 'VLAN Name' | Table -Name 'VLANs' 
                                }
                            }
                            #endregion VLANs

                            #region Server Links
                            $UcsServerLinks = Get-UcsFabricPort -Ucs $UCSM | Where-Object { $_.IfRole -eq 'server' }
                            if ($UcsServerLinks) {
                                Section -Style Heading4 -Name 'Server Links' {
                                    $ServerLinks = foreach ($ServerLink in $UcsServerLinks) {
                                        [PSCustomObject]@{
                                            'Fabric' = $ServerLink.SwitchId
                                            'Port' = "$($ServerLink.SlotId)/$($ServerLink.PortId)"
                                            'MAC' = $ServerLink.Mac
                                            'State' = $ServerLink.AdminState
                                            'Chassis ID' = $ServerLink.ChassisId
                                            'Aggregated Port ID' = $ServerLink.AggrPortId
                                            'If Role' = $ServerLink.IfRole
                                            'If Type' = $ServerLink.IfType
                                            'Network Type' = $ServerLink.Type
                                            'Transport' = $ServerLink.Transport -join ', '
                                            'Speed' = $ServerLink.OperSpeed
                                            'Status' = $ServerLink.OperState
                                        }
                                    }
                                    $ServerLinks | Sort-Object 'Fabric', 'Port' | Table -Name 'Server Links'
                                }
                            }
                            #endregion Server Links

                            #region QoS System Class Section
                            $UcsQosClass = Get-UcsQosClass -Ucs $UCSM | Sort-Object Cos -Descending
                            if ($UcsQosClass) {
                                $UcsFcQosClass = Get-UcsFcQosClass -Ucs $UCSM
                                $UcsBestEffortQosClass = Get-UcsBestEffortQosClass -Ucs $UCSM
                                $UcsQosClass += $UcsBestEffortQosClass
                                $UcsQosClass += $UcsFcQosClass    
                                Section -Style Heading4 -Name 'QoS System Class' {
                                    $QosClasses = foreach ($QosClass in $UcsQosClass) {
                                        [PSCustomObject]@{
                                            'Priority' = $QosClass.Priority
                                            'Enabled' = $QosClass.AdminState
                                            'CoS' = $QosClass.Cos
                                            'Packet Drop' = Switch ($QosClass.Drop) {
                                                'no-drop' { 'disabled' }
                                                'drop' { 'enabled' }
                                            }
                                            'Weight' = $QosClass.Weight
                                            'Weight (%)' = $QosClass.BwPercent
                                            'MTU' = $QosClass.Mtu
                                        }
                                    }
                                    $QosClasses | Table -Name 'QoS System Class'
                                }
                            }
                            #endregion QoS System Class Section

                            #region LAN Pin Groups Section
                            $UcsLanPinGroups = Get-UcsEthernetPinGroup -Ucs $UCSM
                            if ($UcsLanPinGroups) {
                                Section -Style Heading4 -Name 'LAN Pin Groups' {
                                    $LanPinGroups = foreach ($LanPinGroup in $UcsLanPinGroups) {
                                        $LanPinGroupTarget = Get-UcsEthernetPinGroupTarget -EthernetPinGroup $LanPinGroup
                                        [PSCustomObject]@{
                                            'Name' = $LanPinGroup.Name
                                            'Description' = $LanPinGroup.Descr
                                            'Target Ports' = $LanPinGroupTarget.EpDn -join ', '
                                        }
                                    }
                                    $LanPinGroups | Sort-Object 'Name' | Table -Name 'LAN Pin Groups' 
                                }
                            }
                            #endregion LAN Pin Groups Section
                        }
                        #endregion LAN Cloud Section

                        #region LAN Appliances Section
                        $UcsAppliance = Get-UcsAppliance -Ucs $UCSM
                        if ($UcsAppliance) {
                            Section -Style Heading3 -Name 'Appliances' {
                            }
                        }
                        #endregion LAN Appliances Section

                        #region LAN Policies Section
                        Section -Style Heading3 -Name 'Policies' {
                            #region Global Policies Section
                            $UcsFabricLanCloudPolicy = Get-UcsFabricLanCloudPolicy -Ucs $UCSM
                            if ($UcsFabricLanCloudPolicy) {
                                Section -Style Heading4 -Name 'LAN Global Policies' {
                                    $FabricLanCloudPolicy = [PSCustomObject]@{
                                        'Ethernet Mode' = $UcsFabricLanCloudPolicy.Mode
                                        'MAC Address Table Aging' = $UcsFabricLanCloudPolicy.MacAging
                                        'VLAN Port Count Optimization' = $UcsFabricLanCloudPolicy.VlanCompression
                                    }
                                    $FabricLanCloudPolicy | Table -List -Name 'LAN Global Policies' -ColumnWidths 50, 50
                                }
                            }
                            #endregion Global Policies Section

                            <#
                        #region Default vNIC Behavior Policy Section
                        $UcsVnicDefBeh = Get-UcsVnicDefBeh -Ucs $UCSM
                        if ($UcsVnicDefBeh) {
                            Section -Style Heading4 -Name 'Default vNIC Behavior' {
                            }
                        }
                        #endregion Default vNIC Behavior Policy Section

                        #region Flow Control Policy Section
                        $UcsFlowControlPolicy = Get-UcsFlowControlPolicy -Ucs $UCSM
                        if ($UcsFlowControlPolicy) {
                            Section -Style Heading4 -Name 'Flow Control' {
                            }
                        }
                        #endregion Flow Control Policy Section

                        #region Dynamic vNIC Connection Policy Section
                        $UcsDynamicVnicConnPolicy = Get-UcsDynamicVnicConnPolicy -Ucs $UCSM
                        if ($UcsDynamicVnicConnPolicy) {
                            Section -Style Heading4 -Name 'Dynamic vNIC Connection' {
                            }
                        }
                        #endregion Dynamic vNIC Connection Policy Section
                        #>

                            #region LACP Policy Section
                            $UcsFabricLacpPolicy = Get-UcsFabricLacpPolicy -Ucs $UCSM
                            if ($UcsFabricLacpPolicy) {
                                Section -Style Heading4 -Name 'LACP' {
                                    $FabricLacpPolicy = foreach ($LacpPolicy in $UcsFabricLacpPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $LacpPolicy.Name
                                            'Distinguished Name' = $LacpPolicy.Dn
                                            'Owner' = $LacpPolicy.PolicyOwner
                                            'Suspend Individual' = $LacpPolicy.SuspendIndividual
                                            'LACP Rate' = $LacpPolicy.FastTimer
                                        }
                                    }
                                    $FabricLacpPolicy | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'LACP Policies'
                                }
                            }
                            #endregion LACP Policy Section

                            <#
                            #region LAN Connectivity Section
                            Section -Style Heading4 -Name 'LAN Connectivity' {
                            }
                            #endregion LAN Connectivity Section
                            #>

                            #region Link Protocol Policy Section
                            $UcsFabricUdldPolicy = Get-UcsFabricUdldPolicy -Ucs $UCSM
                            if ($UcsFabricUdldPolicy) {
                                Section -Style Heading4 -Name 'Link Protocol' {
                                    $FabricUdldPolicy = [PSCustomObject]@{
                                        'Name' = $UcsFabricUdldPolicy.Name
                                        'Owner' = $UcsFabricUdldPolicy.PolicyOwner
                                        'Message Interval' = $UcsFabricUdldPolicy.MsgInterval
                                        'Recovery Action' = $UcsFabricUdldPolicy.RecoveryAction
                                    }
                                    $FabricUdldPolicy | Sort-Object 'Name', 'Owner' | Table -Name 'Link Protocol Policies' -ColumnWidths 25, 25, 25, 25
                                }
                            }
                            #endregion Link Protocol Policy Section

                            #region Multicast Policy Section
                            $UcsFabricMulticastPolicy = Get-UcsFabricMulticastPolicy -Ucs $UCSM
                            if ($UcsFabricMulticastPolicy) {
                                Section -Style Heading4 -Name 'Multicast' {
                                    $FabricMulticastPolicy = foreach ($MulticastPolicy in $UcsFabricMulticastPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $MulticastPolicy.Name
                                            'Owner' = $MulticastPolicy.PolicyOwner
                                            'IGMP Snooping State' = $MulticastPolicy.SnoopingState
                                            'IGMP Snooping Querier State' = $MulticastPolicy.QuerierState
                                            'FI-A Querier IPv4 Address' = $MulticastPolicy.QuerierIpAddr
                                            'FI-B Querier IPv4 Address' = $MulticastPolicy.QuerierIpAddrPeer
                                        }
                                    }
                                    $FabricMulticastPolicy | Sort-Object 'Name', 'Owner' | Table -Name 'Multicast Policies'
                                }
                            }
                            #endregion Multicast Policy Section

                            #region Network Control Policy Section
                            $UcsNetworkControlPolicy = Get-UcsNetworkControlPolicy
                            if ($UcsNetworkControlPolicy) {
                                Section -Style Heading4 -Name 'Network Control' {
                                    $NetworkControlPolicy = foreach ($NetControlPolicy in $UcsNetworkControlPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $NetControlPolicy.Name
                                            'Owner' = $NetControlPolicy.PolicyOwner
                                            'CDP' = $NetControlPolicy.Cdp
                                            'MAC Register Mode' = $NetControlPolicy.MacRegisterMode
                                            'Action on Uplink Fail' = $NetControlPolicy.UplinkFailAction
                                            'LLDP Transmit' = $NetControlPolicy.LldpTransmit
                                            'LLDP Receive' = $NetControlPolicy.LldpReceive
                                        }
                                    }
                                    $NetworkControlPolicy | Sort-Object 'Name', 'Owner' | Table -Name 'Network Control Policies'
                                }
                            }
                            #endregion Network Control Policy Section

                            <#
                            #region QoS Policy Section
                            $UcsQosPolicy = Get-UcsQosPolicy -Ucs $UCSM
                            if ($UcsQosPolicy) {
                                Section -Style Heading4 -Name 'QoS' {
                                }
                            }
                            #endregion QoS Policy Section

                            #region Threshold Policy Section
                            $UcsThresholdPolicy = Get-UcsThresholdPolicy -Ucs $UCSM
                            if ($UcsThresholdPolicy) {
                                Section -Style Heading4 -Name 'Threshold' {
                                    $ThresholdPolicy = foreach ($ThresholdPol in $UcsThresholdPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $ThresholdPol.Name
                                            'Distingushed Name' = $ThresholdPol.Dn
                                            'Description' = $ThresholdPol.Descr
                                            'Owner' = $ThresholdPol.PolicyOwner
                                            'Default Thresholds Added' = $ThresholdPol.DefaultThresholdsAdded
                                        }
                                    }
                                    $ThresholdPolicy | Sort-Object 'Name','Distingushed Name' | Table -Name 'Threshold Policies'
                                }
                            }
                            #endregion Threshold Policy Section

                            #region VMQ Connection Policy Section
                            Section -Style Heading4 -Name 'VMQ Connection' {
                            }
                            #endregion VMQ Connection Policy Section
                            #>

                            #region usNIC Connection Policy Section
                            $UcsVnicUsnicConPolicy = Get-UcsVnicUsnicConPolicy -Ucs $UCSM
                            if ($UcsVnicUsnicConPolicy) {
                                Section -Style Heading4 -Name 'usNIC Connection' {
                                    $VnicUsnicConPolicy = foreach ($UsnicConPolicy in $UcsVnicUsnicConPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $UsnicConPolicy.Name
                                            'Owner' = $UsnicConPolicy.PolicyOwner
                                            'Description' = $UsnicConPolicy.Descr
                                            'Number of usNICs' = $UsnicConPolicy.UsnicCount
                                            'Adapter Policy' = $UsnicConPolicy.AdaptorProfileName
                                        }
                                    }
                                    $VnicUsnicConPolicy | Sort-Object 'Name', 'Owner' | Table -Name 'usNIC Connection Policies'
                                }
                            }
                            #endregion usNIC Connection Policy Section

                            #region vNIC Template Section
                            $UcsVnicTemplate = Get-UcsVnicTemplate -Ucs $UCSM
                            if ($UcsVnicTemplate) {
                                Section -Style Heading4 -Name 'vNIC Templates' {
                                    $VnicTemplate = foreach ($Vnic in $UcsVnicTemplate) {
                                        [PSCustomObject]@{
                                            'Name' = $Vnic.Name
                                            'Description' = $Vnic.Descr
                                            'Owner' = $Vnic.PolicyOwner
                                            'Fabric' = Switch ($Vnic.SwitchId) {
                                                'A-B' { 'enable failover' }
                                                'B-A' { 'enable failover' }
                                                default { $Vnic.SwitchId }
                                            }
                                            'Redundancy Type' = $Vnic.RedundancyPairType
                                            'Target' = $Vnic.Target -join ', '
                                            'Template Type' = $Vnic.TemplType
                                            'CDN Source' = $Vnic.CdnSource
                                            'MTU' = $Vnic.Mtu
                                            'MAC Pool' = $Vnic.IdentPoolName
                                            'QoS Policy' = Switch ($Vnic.QosPolicyName) {
                                                '' { 'not set' }
                                                default { $Vnic.QosPolicyName }
                                            }
                                            'Network Control Policy' = Switch ($Vnic.NwCtrlPolicyName) {
                                                '' { 'not set' }
                                                default { $Vnic.NwCtrlPolicyName }
                                            }
                                            'Pin Group' = $Vnic.PinToGroupName
                                            'Stats Threshold Policy' = $Vnic.StatsPolicyName
                                            'VLANs' = ($Vnic | Get-UcsChild).Name -join ', '
                                        }
                                    }
                                    $VnicTemplate | Sort-Object 'Name' | Table -List -Name 'vNIC Templates' -ColumnWidths 50, 50
                                }
                            }
                            #endregion vNIC Template Section
                        }
                        #endregion LAN Policies Section

                        #region Pools Section
                        Section -Style Heading3 -Name 'Pools' {
                            #region IP Pools Section
                            $UcsIpPool = Get-UcsIpPool -Ucs $UCSM | Where-Object { $_.Size -gt 0 }
                            if ($UcsIpPool) {
                                Section -Style Heading4 -Name 'IP Pools' {
                                    $IpPools = foreach ($IpPool in $UcsIpPool) {
                                        $UcsIpPoolBlock = Get-UcsIpPoolBlock -IpPool $IpPool
                                        [PSCustomObject]@{
                                            'Name' = $IpPool.Name
                                            'Owner' = $IpPool.PolicyOwner
                                            'Description' = $IpPool.Descr
                                            #'GUID' = $IpPool.Guid
                                            'Size' = $IpPool.Size
                                            'Assigned' = $IpPool.Assigned
                                            'Assignment Order' = $IpPool.AssignmentOrder
                                            'IP Blocks' = ($UcsIpPoolBlock | Sort-Object From | ForEach-Object { "$($_.From) - $($_.To)" }) -join [Environment]::NewLine
                                        }
                                    }
                                    $IpPools | Sort-Object 'Name' | Table -Name 'IP Pools' 
                                }
                            }
                            #endregion IP Pools Section

                            #region IP Pool Addresses
                            $UcsIpPoolAddr = Get-UcsIpPoolAddr -Ucs $UCSM
                            if ($UcsIpPoolAddr) {
                                Section -Style Heading4 -Name 'IP Pool Addresses' {
                                    $IpPoolAddr = foreach ($IpAddr in $UcsIpPoolAddr) {
                                        [PSCustomObject]@{
                                            'IP Address' = $IpAddr.Id
                                            'Owner' = $IpAddr.Owner
                                            'Assigned' = $IpAddr.Assigned
                                            'Assigned To' = $IpAddr.AssignedToDn
                                        }
                                    }
                                    $IpPoolAddr | Sort-Object 'IP Address' | Table -Name 'IP Pool Addresses'
                                }
                            }
                            #endregion IP Pool Addresses
                        
                            #region MAC Pools Section
                            $UcsMacPool = Get-UcsMacPool -Ucs $UCSM | Where-Object { $_.Size -gt 0 }
                            if ($UcsMacPool) {
                                Section -Style Heading4 -Name 'MAC Pools' {
                                    $MacPools = foreach ($MacPool in $UcsMacPool) {
                                        $UcsMacPoolBlock = Get-UcsMacMemberBlock -MacPool $MacPool
                                        [PSCustomObject]@{
                                            'Name' = $MacPool.Name
                                            'Owner' = $MacPool.PolicyOwner
                                            'Description' = $MacPool.Descr
                                            'Size' = $MacPool.Size
                                            'Assigned' = $MacPool.Assigned
                                            'Assignment Order' = $MacPool.AssignmentOrder
                                            'MAC Blocks' = ($UcsMacPoolBlock | Sort-Object From | ForEach-Object { "$($_.From) - $($_.To)" }) -join [Environment]::NewLine
                                        }
                                    }
                                    $MacPools | Sort-Object 'Name' | Table -Name 'MAC Pools'
                                }
                            }
                            #endregion MAC Pools Section

                            #region MAC Pool Addresses
                            $UcsMacPoolAddr = Get-UcsMacPoolAddr -Ucs $UCSM | Where-Object { $_.Assigned -eq 'yes' }
                            if ($UcsMacPoolAddr) {
                                Section -Style Heading4 -Name 'MAC Pool Addresses' {
                                    $MacPoolAddr = foreach ($MacAddr in $UcsMacPoolAddr) {
                                        [PSCustomObject]@{
                                            'MAC Address' = $MacAddr.Id
                                            'Owner' = $MacAddr.Owner
                                            'Assigned' = $MacAddr.Assigned
                                            'Assigned To' = $MacAddr.AssignedToDn
                                        }
                                    }
                                    $MacPoolAddr | Sort-Object 'MAC Address' | Table -Name 'MAC Pool Addresses'
                                }
                            }
                            #endregion MAC Pool Addresses
                        }
                        #endregion Pools Section
                    }
                }
                #endregion LAN Section

                #region SAN Section
                if ($Section.SAN) {
                    Section -Style Heading2 -Name 'SAN' {
                        #region SAN Cloud Section
                        Section -Style Heading3 -Name 'SAN Cloud' {

                            $UcsSanCloud = Get-UcsSanCloud -Ucs $UCSM
                            if ($UcsSanCloud) {
                                $SanCloud = [PSCustomObject]@{
                                    'UCS' = $UcsSanCloud.Ucs
                                    'Mode' = $UcsSanCloud.Mode
                                } 
                                $SanCloud | Table -Name 'SAN Cloud' -ColumnWidths 50, 50
                            }

                            #region FC Port Channels
                            $UcsFcUplinkPortChannel = Get-UcsFcUplinkPortChannel -Ucs $UCSM
                            if ($UcsFcUplinkPortChannel) {
                                Section -Style Heading4 -Name 'FC Port Channels' {
                                    $FcUplinkPortChannel = foreach ($FcPortChannel in $UcsFcUplinkPortChannel) {
                                        [PSCustomObject]@{
                                            'Fabric' = $FcPortChannel.SwitchId
                                            'Port ID' = $FcPortChannel.PortId
                                            'Name' = $FcPortChannel.Name
                                            'Description' = $FcPortChannel.Descr
                                            'If Type' = $FcPortChannel.IfType
                                            'Transport' = $FcPortChannel.Transport
                                            'Admin Speed' = $FcPortChannel.AdminSpeed
                                            'Operational Speed Gbps' = $FcPortChannel.OperSpeed
                                            #'Ports'
                                            'Status' = $FcPortChannel.OperState
                                        }
                                    }
                                    $FcUplinkPortChannel | Sort-Object 'Fabric' | Table -Name 'FC Port Channels'
                                }
                            }
                            #endregion FC Port Channels

                            #region FC Port Channels
                            $UcsFabricFcoeSanPc = Get-UcsFabricFcoeSanPc -Ucs $UCSM
                            if ($UcsFabricFcoeSanPc) {
                                Section -Style Heading4 -Name 'FCoE Port Channels' {
                                }
                            }
                            #endregion FC Port Channels

                            #region Uplink FC Interfaces
                            $UcsFiFcPort = Get-UcsFiFcPort -Ucs $UCSM
                            if ($UcsFiFcPort) {
                                Section -Style Heading4 -Name 'Uplink FC Interfaces' {
                                    $FiFcPorts = foreach ($FcPort in $UcsFiFcPort) {
                                        [PSCustomObject]@{
                                            'Fabric' = $FcPort.SwitchId
                                            'Port' = "$($FcPort.SlotId)/$($FcPort.PortId)"
                                            'WWN' = $FcPort.wwn
                                            'Unified Port' = $FcPort.UnifiedPort
                                            'Port Channel Member' = $FcPort.IsPortChannelMember
                                            'State' = $FcPort.AdminState
                                            'If Role' = $FcPort.IfRole
                                            'If Type' = $FcPort.IfType
                                            'Network Type' = $FcPort.Type
                                            'Transport' = $FcPort.Transport -join ', '
                                            'Speed' = $FcPort.OperSpeed
                                            'Status' = $FcPort.OperState
                                        }
                                    }
                                    $FiFcPorts | Sort-Object 'Fabric', 'Port' | Table -Name 'Uplink FC Interfaces'
                                }
                            }
                            #endregion Uplink FC Interfaces

                            #region Uplink FCoE Interfaces
                            $UcsFabricPort = Get-UcsFabricPort -Ucs $UCSM | Where-Object { $_.IfRole -eq 'fcoe-uplink' }
                            if ($UcsFabricPort) {
                                Section -Style Heading4 -Name 'Uplink FCoE Interfaces' {
                                }
                            }
                            #endregion Uplink FCoE Interfaces

                            #region VSANs
                            $UcsVsan = Get-UcsVsan -Ucs $UCSM
                            if ($UcsVsan) {
                                Section -Style Heading4 -Name 'VSANs' {
                                    $Vsans = foreach ($Vsan in $UcsVsan) {
                                        [PSCustomObject]@{
                                            'Name' = $Vsan.Name
                                            'ID' = $Vsan.Id
                                            'Fabric' = $Vsan.SwitchId
                                            'If Type' = $Vsan.IfType
                                            'If Role' = $Vsan.IfRole
                                            'Transport' = $Vsan.Transport
                                            'FCoE VLAN' = $Vsan.FcoeVlan
                                            'FC Zoning' = $Vsan.ZoningState
                                            'Default FC Zoning' = $Vsan.DefaultZoning
                                            'FC Zone Sharing Mode' = $Vsan.FcZoneSharingMode
                                            'Status' = $Vsan.OperState
                                        }
                                    }
                                    $Vsans | Sort-Object 'Fabric', 'Name', 'ID' | Table -Name 'VSANs'
                                }
                            }
                            #endregion VSANs
                        }
                        #endregion SAN Cloud Section

                        <#
                    ##TODO
                    Section -Style Heading3 -Name 'Storage Cloud' {
                        Section -Style Heading4 -Name 'Storage FC Interfaces' {
                        }

                        Section -Style Heading4 -Name 'Storage FCoE Interfaces' {
                        }

                        Section -Style Heading4 -Name 'FC Zone Profiles' {
                        }
                    }
                    #>

                        #region SAN Policies Section
                        Section -Style Heading3 -Name 'Policies' {
                            #region Default vHBA Behavior
                            $UcsVnicVhbaBehPolicy = Get-UcsVnicVhbaBehPolicy -Ucs $UCSM
                            if ($UcsVnicVhbaBehPolicy) {
                                Section -Style Heading4 -Name 'Default vHBA Behavior' {
                                    $VhbaBehPolicy = foreach ($VhbaPolicy in $UcsVnicVhbaBehPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $VhbaPolicy.Name
                                            'Distinguished Name' = $VhbaPolicy.Dn
                                            'Description' = $VhbaPolicy.Descr
                                            'Type' = $VhbaPolicy.
                                            'Action' = $VhbaPolicy.Action
                                        }
                                    }
                                    $VhbaBehPolicy | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Default vHBA Behavior'
                                }
                            }
                            #endregion Default vHBA Behavior

                            #region Fibre Channel Adapter Policy
                            $UcsFcAdapterPolicy = Get-UcsFcAdapterPolicy -Ucs $UCSM
                            if ($UcsFcAdapterPolicy) {
                                Section -Style Heading4 -Name 'Fibre Channel Adapter Policies' {
                                    $FcAdapterPolicies = foreach ($FcAdapterPolicy in $UcsFcAdapterPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $FcAdapterPolicy.Name
                                            'Distinguished Name' = $FcAdapterPolicy.Dn
                                            'Description' = $FcAdapterPolicy.Descr
                                            'Owner' = $FcAdapterPolicy.PolicyOwner
                                        }
                                    }
                                    $FcAdapterPolicies | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'Fibre Channel Adapter Policies'
                                }
                            }
                            #endregion Fibre Channel Adapter Policy
                            
                            #region LACP Policy Section
                            $UcsFabricLacpPolicy = Get-UcsFabricLacpPolicy -Ucs $UCSM
                            if ($UcsFabricLacpPolicy) {
                                Section -Style Heading4 -Name 'LACP' {
                                    $FabricLacpPolicy = foreach ($LacpPolicy in $UcsFabricLacpPolicy) {
                                        [PSCustomObject]@{
                                            'Name' = $LacpPolicy.Name
                                            'Distinguished Name' = $LacpPolicy.Dn
                                            'Owner' = $LacpPolicy.PolicyOwner
                                            'Suspend Individual' = $LacpPolicy.SuspendIndividual
                                            'LACP Rate' = $LacpPolicy.FastTimer
                                        }
                                    }
                                    $FabricLacpPolicy | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'LACP Policies'
                                }
                            }
                            #endregion LACP Policy Section

                            <#
                            Section -Style Heading4 -Name 'SAN Connectivity' {
                            }

                            Section -Style Heading4 -Name 'Storage Connection' {
                            }

                            Section -Style Heading4 -Name 'Threshold' {
                            }
                            #>

                            #region vHBA Templates
                            $UcsVhbaTemplate = Get-UcsVhbaTemplate
                            if ($UcsVhbaTemplate) {
                                Section -Style Heading4 -Name 'vHBA Templates' {
                                    $VhbaTemplates = foreach ($VhbaTemplate in $UcsVhbaTemplate) {
                                        [PSCustomObject]@{
                                            'Name' = $VhbaTemplate.Name
                                            'Description' = $VhbaTemplate.Descr
                                            'Owner' = $VhbaTemplate.PolicyOwner
                                            'Fabric' = $VhbaTemplate.SwitchId
                                            'Redundancy' = $VhbaTemplate.RedundancyPairType
                                            'Target' = $VhbaTemplate.Target
                                            'Template Type' = $VhbaTemplate.TemplType
                                            'Max Data Field Size' = $VhbaTemplate.MaxDataFieldSize
                                            'WWPN Pool' = $VhbaTemplate.IdentPoolName
                                            'QoS Policy' = $VhbaTemplate.NwCtrlPolicyName
                                            'Pin Group' = $VhbaTemplate.PinToGroupName
                                            'Stats Threshold Policy' = $VhbaTemplate.StatsPolicyName
                                        }
                                    }
                                    $VhbaTemplates | Sort-Object 'Name' | Table -List -Name 'vHBA Templates' -ColumnWidths 50, 50
                                
                                    Section -Style Heading4 -Name 'vHBA Interfaces' {
                                        $UcsVhbaInterface = $VhbaTemplate | Get-UcsChild
                                        $VhbaInterfaces = foreach ($VhbaIf in $UcsVhbaInterface) {
                                            [PSCustomObject]@{
                                                'Name' = $VhbaIf.Name
                                                'WWPN' = $VhbaIf.Initiator
                                                'Owner' = $VhbaIf.Owner
                                                'Type' = $VhbaIf.Type
                                            }
                                        }
                                        $VhbaInterfaces | Sort-Object 'Name ' | Table -Name 'vHBA Interfaces'
                                    }
                                }
                            }
                            #endregion vHBA Templates
                        }
                        #endregion SAN Policies Section

                        #region SAN Pools
                        Section -Style Heading3 -Name 'Pools' {
                            #region IQN Pools
                            $UcsIqnPoolPool = Get-UcsIqnPoolPool -Ucs $UCSM | Where-Object { $_.Size -gt 0 }
                            if ($UcsIqnPoolPool) {
                                Section -Style Heading4 -Name 'IQN Pools' {
                                    $IqnPools = foreach ($IqnPool in $UcsIqnPoolPool) {
                                        $UcsIqnPoolBlock = Get-UcsIqnPoolBlock -IqnPoolPool $IqnPool 
                                        [PSCustomObject]@{
                                            'Name' = $IqnPool.Name
                                            'Distinguished Name' = $IqnPool.Dn
                                            'Description' = $IqnPool.Descr
                                            'Prefix' = $IqnPool.Prefix
                                            'Size' = $IqnPool.Size
                                            'Assigned' = $IqnPool.Assigned
                                            'Assignment Order' = $IqnPool.AssignmentOrder
                                            'IQN Blocks' = ($UcsIqnPoolBlock | Sort-Object Suffix | ForEach-Object { "$($_.Suffix):$($_.From) - $($_.Suffix):$($_.To)" }) -join [Environment]::NewLine
                                        }
                                    }
                                    $IqnPools | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'IQN Pools'
                                }
                            }
                            #end region IQN Pools

                            #region WWNN Pools
                            $UcsWwnnPool = Get-UcsWwnPool -Ucs $UCSM | Where-Object { $_.Purpose -eq 'node-wwn-assignment' -and $_.Size -gt 0 }
                            if ($UcsWwnnPool) {
                                Section -Style Heading4 -Name 'WWNN Pools' {
                                    $WwnnPools = foreach ($WwnnPool in $UcsWwnnPool) {
                                        $UcsWwnnMemberBlock = Get-UcsWwnMemberBlock -WwnPool $WwnnPool 
                                        [PSCustomObject]@{
                                            'Name' = $WwnnPool.Name
                                            'Distinguished Name' = $WwnnPool.Dn
                                            'Description' = $WwnnPool.Descr
                                            'Size' = $WwnnPool.Size
                                            'Assigned' = $WwnnPool.Assigned
                                            'Assignment Order' = $WwnnPool.AssignmentOrder
                                            'WWN Initiator Blocks' = ($UcsWwnnMemberBlock | Sort-Object From | ForEach-Object { "$($_.From) - $($_.To)" }) -join [Environment]::NewLine
                                        }
                                    }
                                    $WwnnPools | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'WWNN Pools'
                                }
                            }
                            #endregion WWNN Pools

                            #region WWPN Pools
                            $UcsWwpnPool = Get-UcsWwnPool -Ucs $UCSM | Where-Object { $_.Purpose -eq 'port-wwn-assignment' -and $_.Szie -gt 0 }
                            if ($UcsWwpnPool) {
                                Section -Style Heading4 -Name 'WWPN Pools' {
                                    $WwpnPools = foreach ($WwpnPool in $UcsWwpnPool) {
                                        $UcsWwpnMemberBlock = Get-UcsWwnMemberBlock -WwnPool $WwpnPool 
                                        [PSCustomObject]@{
                                            'Name' = $WwpnPool.Name
                                            'Distinguished Name' = $WwpnPool.Dn
                                            'Description' = $WwpnPool.Descr
                                            'Size' = $WwpnPool.Size
                                            'Assigned' = $WwpnPool.Assigned
                                            'Assignment Order' = $WwpnPool.AssignmentOrder
                                            'WWN Initiator Blocks' = ($UcsWwpnMemberBlock | Sort-Object From | ForEach-Object { "$($_.From) - $($_.To)" }) -join [Environment]::NewLine
                                        }
                                    }
                                    $WwpnPools | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'WWPN Pools'
                                }
                            }
                            #endregion WWPN Pools

                            #region WWxN Pools
                            $UcsWwxnPool = Get-UcsWwnPool -Ucs $UCSM | Where-Object { $_.Purpose -eq 'node-and-port-wwn-assignment' -and $_.Size -gt 0 }
                            if ($UcsWwxnPool) {
                                Section -Style Heading4 -Name 'WWxN Pools' {
                                    $WwxnPools = foreach ($WwxnPool in $UcsWwxnPool) {
                                        $UcsWwxnMemberBlock = Get-UcsWwnMemberBlock -WwnPool $WwxnPool 
                                        [PSCustomObject]@{
                                            'Name' = $WwxnPool.Name
                                            'Distinguished Name' = $WwxnPool.Dn
                                            'Description' = $WwxnPool.Descr
                                            'Size' = $WwxnPool.Size
                                            'Assigned' = $WwxnPool.Assigned
                                            'Assignment Order' = $WwxnPool.AssignmentOrder
                                            'WWN Blocks' = ($UcsWwxnMemberBlock | Sort-Object From | ForEach-Object { "$($_.From) - $($_.To)" }) -join [Environment]::NewLine
                                        }
                                    }
                                    $WwxnPools | Sort-Object 'Name', 'Distinguished Name' | Table -Name 'WWxN Pools'
                                }
                            }
                            #endregion WWxN Pools
                        }
                        #endregion SAN Pools


                        <#
                    Section -Style Heading2 -Name 'SAN Cloud' {
                        Section -Style Heading3 -Name 'Fabric Interconnect Fibre Channel Switching Mode' {
                            $UcsSanCloud = Get-UcsSanCloud | Select-Object Rn, Mode
                            $UcsSanCloud | Table -Name 'Fabric Interconnect Fibre Channel Switching Mode' 
                        }

                        Section -Style Heading3 -Name 'Fabric Interconnect FC Uplink Ports' {
                            $UcsFiFcPort = Get-UcsFiFcPort | Select-Object EpDn, SwitchId, SlotId, PortId, LicState, Mode, OperSpeed, OperState, wwn | Sort-Object -descending  | where-object {$_.OperState -ne 'sfp-not-present'}
                            $UcsFiFcPort | Table -Name 'Fabric Interconnect FC Uplink Ports' 
                        }

                        Section -Style Heading3 -Name 'Fabric Interconnect FC Uplink Port Channels' {
                            $UcsFcUplinkPortChannel = Get-UcsFcUplinkPortChannel | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Name, OperSpeed, OperState, Transport
                            $UcsFcUplinkPortChannel | Table -Name 'Fabric Interconnect FC Uplink Port Channels' 
                        }

                        Section -Style Heading3 -Name 'Fabric Interconnect FCoE Uplink Ports' {
                            $UcsFabricPort = Get-UcsFabricPort | Where-Object {$_.IfRole -eq 'fcoe-uplink'} | Select-Object IfRole, EpDn, LicState, OperState, OperSpeed
                            $UcsFabricPort | Table -Name 'Fabric Interconnect FCoE Uplink Ports' 
                        }

                        Section -Style Heading3 -Name 'Fabric Interconnect FCoE Uplink Port Channels' {
                            $UcsFabricFcoeSanPc = Get-UcsFabricFcoeSanPc | Select-Object @{L = 'Distinguished Name'; E = {$_.Dn}}, Name, FcoeState, OperState, Transport, Type
                            $UcsFabricFcoeSanPc | Table -Name 'Fabric Interconnect FCoE Uplink Port Channels' 
                        }

                    }
                    #>
                    }
                }
                #endregion SAN Section
            
                #region VM Section
                ##TODO
                <#
            if ($Section.VM) {
                Section -Style Heading2 -Name 'VM' {
                    Section -Style Heading3 -Name 'Clusters' {
                    }

                    Section -Style Heading3 -Name 'Fabric Network Sets' {
                    }

                    Section -Style Heading3 -Name 'Port Profiles' {
                    }

                    Section -Style Heading3 -Name 'VM Networks' {
                    }

                    Section -Style Heading3 -Name 'Microsoft' {
                    }

                    Section -Style Heading3 -Name 'VMware' {
                    }
                }
            }
            #>
                #endregion VM Section

                #region Storage Section
                ##TODO
                <#
            if ($Section.Storage) {
                Section -Style Heading2 -Name 'Storage' {
                    Section -Style Heading3 -Name 'Storage Profiles' {
                    }

                    Section -Style Heading3 -Name 'Storage Policies' {
                    }
                }
            }
            #>
                #endregion Storage Section

                #region Chassis Section
                ##TODO
                <#
            if ($Section.Chassis) {
                Section -Style Heading2 -Name 'Chassis' {
                    Section -Style Heading3 -Name 'Chassis Profiles' {
                    }

                    Section -Style Heading3 -Name 'Chassis Profile Templates' {
                    }

                    Section -Style Heading3 -Name 'Policies' {
                    }
                }
            }
            #>
                #endregion Chassis Section

                #region Admin Section
                if ($Section.Admin) {
                    Section -Style Heading2 -Name 'Admin' {
                        #region User Management
                        Section -Style Heading3 -Name 'User Management' {
                            #region Authentication Section
                            Section -Style Heading4 -Name 'Authentication' {
                                #region Native Authentication Section
                                $UcsNativeAuth = Get-UcsNativeAuth -Ucs $UCSM
                                if ($UcsNativeAuth) {
                                    $UcsDefaultAuth = Get-UcsDefaultAuth -Ucs $UCSM
                                    Section -Style Heading4 -Name 'Native Authentication' {
                                        $NativeAuth = [PSCustomObject]@{
                                            'Default Authentication Realm' = $UcsDefaultAuth.Realm
                                            'Web Session Refresh Period (sec)' = $UcsDefaultAuth.RefreshPeriod
                                            'Web Session Timeout (sec)' = $UcsDefaultAuth.SessionTimeout
                                            'Use Two Factor Authentication' = $UcsDefaultAuth.Use2Factor
                                            'Console Authentication Realm' = $UcsNativeAuth.ConLogin
                                            'Role Policy for Remote Users' = $UcsNativeAuth.DefRolePolicy
                                        }
                                        $NativeAuth | Table -List -Name 'Native Authentication' -ColumnWidths 50, 50    
                                    }
                                }
                                #endregion Native Authentication Section

                                #region Domain Authentication Section
                                $UcsAuthDomains = Get-UcsAuthDomain -Ucs $UCSM
                                if ($UcsAuthDomains) {
                                    Section -Style Heading4 -Name 'Authentication Domains' {
                                        $AuthDomain = foreach ($UcsAuthDomain in $UcsAuthDomains) {
                                            $DefaultAuthDomain = Get-UcsAuthDomainDefaultAuth -AuthDomain $UcsAuthDomain
                                            [PSCustomObject]@{
                                                'Name' = $UcsAuthDomain.Name
                                                'Web Session Refresh Period (sec)' = $UcsAuthDomain.RefreshPeriod
                                                'Web Session Timeout (sec)' = $UcsAuthDomain.SessionTimeout
                                                'Realm' = $DefaultAuthDomain.Realm
                                                'Provider Group' = $DefaultAuthDomain.ProviderGroup
                                            }
                                        }
                                        $AuthDomain | Table -List -Name 'Native Authentication' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion Domain Authentication Section
                            }
                            #endregion Authentication Section

                            #region LDAP Section
                            $UcsLdapGlobalConfig = Get-UcsLdapGlobalConfig -Ucs $UCSM
                            if ($UcsLdapGlobalConfig) {
                                Section -Style Heading4 -Name 'LDAP' {
                                    #region LDAP Properties
                                    $LdapGlobalConfig = [PSCustomObject]@{
                                        'Timeout' = $UcsLdapGlobalConfig.Timeout
                                        'Attribute' = $UcsLdapGlobalConfig.Attribute
                                        'Base DN' = $UcsLdapGlobalConfig.BaseDn
                                        'Filter' = $UcsLdapGlobalConfig.Filter
                                    }
                                    $LdapGlobalConfig | Table -List -Name 'LDAP Global Config' -ColumnWidths 50, 50
                                    #endregion LDAP Properties

                                    #region LDAP Group Map Section
                                    $UcsLdapGroupMaps = Get-UcsLdapGroupMap -Ucs $UCSM
                                    if ($UcsLdapGroupMaps) {
                                        Section -Style Heading4 -Name 'LDAP Group Maps' {
                                            $LdapGroupMap = foreach ($UcsLdapGroupMap in $UcsLdapGroupMaps) {
                                                $UcsUserRole = Get-UcsUserRole -LdapGroupMap $UcsLdapGroupMap | Sort-Object
                                                $UcsUserLocale = Get-UcsUserLocale -LdapGroupMap $UcsLdapGroupMap | Sort-Object
                                                [PSCustomObject]@{
                                                    'Name' = $UcsLdapGroupMap.Name
                                                    'Roles' = ($UcsUserRole).Name -join ', '
                                                    'Locales' = ($UcsUserLocale).Name -join ', '
                                                }
                                            }
                                            $LdapGroupMap | Table -Name 'LDAP Group Map' -ColumnWidths 50, 25, 25
                                        }
                                    }
                                    #endregion LDAP Group Map Section

                                    #region LDAP Provider Groups Section
                                    $UcsLdapProviderGroups = Get-UcsProviderGroup -LdapGlobalConfig $UcsLdapGlobalConfig
                                    if ($UcsLdapProviderGroups) {
                                        Section -Style Heading4 -Name 'LDAP Provider Groups' {
                                            $LdapProviderGroups = foreach ($LdapProviderGroup in $UcsLdapProviderGroups) {
                                                $LdapProviderRef = Get-UcsProviderReference -ProviderGroup $LdapProviderGroup | Sort-Object Order
                                                [PSCustomObject]@{
                                                    'Name' = $LdapProviderGroup.Name
                                                    'LDAP Providers' = ($LdapProviderRef).Name -join ', '
                                                }
                                            }
                                            $LdapProviderGroups | Table -Name 'LDAP Provider Groups' -ColumnWidths 50, 50
                                        }
                                    }
                                    #endregion LDAP Provider Groups Section

                                    #region LDAP Providers Section
                                    $UcsLdapProviders = Get-UcsLdapProvider -Ucs $UCSM | Sort-Object Order
                                    if ($UcsLdapProviders) {
                                        Section -Style Heading4 -Name 'LDAP Providers' {
                                            $LdapProvider = foreach ($UcsLdapProvider in $UcsLdapProviders) {
                                                $UcsLdapGroupRule = Get-UcsLdapGroupRule -LdapProvider $UcsLdapProvider
                                                [PSCustomObject]@{
                                                    'Hostname or IP Address' = $UcsLdapProvider.Name
                                                    'Order' = $UcsLdapProvider.Order
                                                    'Bind DN' = $UcsLdapProvider.RootDn
                                                    'Base DN' = $UcsLdapProvider.BaseDn
                                                    'Port' = $UcsLdapProvider.Port
                                                    'Enable SSL' = $UcsLdapProvider.EnableSsl
                                                    'Filter' = $UcsLdapProvider.Filter
                                                    'Attribute' = $UcsLdapProvider.Attribute
                                                    'Timeout' = $UcsLdapProvider.Timeout
                                                    'Vendor' = $UcsLdapProvider.Vendor
                                                    'Group Authorization' = $UcsLdapGroupRule.Authorization
                                                    'Group Recursion' = $UcsLdapGroupRule.Traversal
                                                    'Target Attribute' = $UcsLdapGroupRule.TargetAttr
                                                    'Use Primary Group' = $UcsLdapGroupRule.UsePrimaryGroup
                                                }
                                            }
                                            $LdapProvider | Table -List -Name "$($UcsLdapProvider.Name) LDAP Provider" -ColumnWidths 50, 50
                                        }
                                    }
                                    #endregion LDAP Providers Section
                                }
                            }
                            #endregion LDAP Section

                            #region RADIUS Section
                            $UcsRadiusGlobalConfig = Get-UcsRadiusGlobalConfig -Ucs $UCSM
                            if ($UcsRadiusGlobalConfig) {
                                Section -Style Heading4 -Name 'RADIUS' {
                                    #region RADIUS Properties
                                    $RadiusGlobalConfig = [PSCustomObject]@{
                                        'Timeout' = $UcsRadiusGlobalConfig.Timeout
                                        'Retries' = $UcsRadiusGlobalConfig.Retries
                                    }
                                    $RadiusGlobalConfig | Table -List -Name 'RADIUS Global Config' -ColumnWidths 50, 50
                                    #endregion RADIUS Properties

                                    #region RADIUS Provider Groups Section
                                    $UcsRadiusProviderGroups = Get-UcsProviderGroup -RadiusGlobalConfig $UcsRadiusGlobalConfig
                                    if ($UcsRadiusProviderGroups) {
                                        Section -Style Heading4 -Name 'RADIUS Provider Groups' {
                                        }
                                    }
                                    #endregion RADIUS Provider Groups Section

                                    #region RADIUS Providers Section
                                    $UcsRadiusProviders = Get-UcsRadiusProvider -Ucs $UCSM
                                    if ($UcsRadiusProviders) {
                                        Section -Style Heading4 -Name 'RADIUS Providers' {
                                        }
                                    }
                                    #endregion RADIUS Providers Section
                                }
                            }
                            #endregion RADIUS Section

                            #region TACACS Section
                            $UcsTacacsGlobalConfig = Get-UcsTacacsGlobalConfig -Ucs $UCSM
                            if ($UcsTacacsGlobalConfig) {
                                Section -Style Heading4 -Name 'TACACS+' {
                                    #region TACACS Properties
                                    $TacacsGlobalConfig = [PSCustomObject]@{
                                        'Timeout' = $UcsTacacsGlobalConfig.Timeout
                                        'Retries' = $UcsTacacsGlobalConfig.Retries
                                    }
                                    $TacacsGlobalConfig | Table -List -Name 'TACACS+ Global Config' -ColumnWidths 50, 50
                                    #endregion TACACS Properties

                                    #region TACACS Provider Groups Section
                                    $UcsTacacsProviderGroups = Get-UcsProviderGroup -TacacsGlobalConfig $UcsTacacsGlobalConfig
                                    if ($UcsTacacsProviderGroups) {
                                        Section -Style Heading4 -Name 'TACACS+ Provider Groups' {
                                        }
                                    }
                                    #endregion TACACS Provider Groups Section

                                    #region TACACS Providers Section
                                    $UcsTacacsProviders = Get-UcsTacacsProvider -Ucs $UCSM
                                    if ($UcsTacacsProviders) {
                                        Section -Style Heading4 -Name 'TACACS+ Providers' {
                                        }
                                    }
                                    #endregion TACACS Providers Section
                                }
                            }
                            #endregion TACACS Section

                            #region Locales Section
                            $UcsLocale = Get-UcsLocale -Ucs $UCSM
                            if ($UcsLocale) {
                                Section -Style Heading4 -Name 'Locales' {
                                }
                            }
                            #endregion Locales Section

                            #region Roles Section
                            $UcsRoles = Get-UcsRole -Ucs $UCSM
                            if ($UcsRoles) {
                                Section -Style Heading4 -Name 'Roles' {
                                    $Roles = foreach ($UcsRole in $UcsRoles) {
                                        [PSCustomObject]@{
                                            'Role' = $UcsRole.Name
                                            'Privileges' = ($UcsRole.Priv | Sort-Object) -join ', '
                                        }
                                    } 
                                    $Roles | Sort-Object 'Role' | Table -Name 'Roles' -ColumnWidths 50, 50
                                }
                            }
                            #endregion Roles Section
                        }
                        #endregion User Management
                    
                        #region Key Management
                        #TODO
                        #endregion Key Management

                        #region Communication Management Section
                        Section -Style Heading3 -Name 'Communication Management' {
                            #region Communication Services Section
                            Section -Style Heading4 -Name 'Communication Services' {    
                                #region Web Session Limits Section
                                $UcsWebSessionLimit = Get-UcsWebSessionLimit -Ucs $UCSM
                                if ($UcsWebSessionLimit) {
                                    Section -Style Heading4 -Name 'Web Session Limits' {
                                        $UcsWebSessionLimits = [PSCustomObject]@{
                                            'Maximum Sessions Per User' = $UcsWebSessionLimit.SessionsPerUser
                                            'Maximum Sessions' = $UcsWebSessionLimit.TotalSessions
                                            'Maximum Event Interval (in seconds)' = $UcsWebSessionLimit.MaxEventInterval
                                        }
                                        $UcsWebSessionLimits | Table -List -Name 'Web Session Limits' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion Web Session Limits Section

                                #region Shell Session Limits Section
                                $UcsShellSessionLimit = Get-UcsShellSvcLimits -Ucs $UCSM
                                if ($UcsShellSessionLimit) {
                                    Section -Style Heading4 -Name 'Shell Session Limits' {
                                        $UcsShellSessionLimits = [PSCustomObject]@{
                                            'Maximum Sessions Per User' = $UcsShellSessionLimit.SessionsPerUser
                                            'Maximum Sessions' = $UcsShellSessionLimit.TotalSessions                            
                                        }
                                        $UcsShellSessionLimits | Table -List -Name 'Web Session Limits' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion Shell Session Limits Section

                                #region CIMC Web Service Section
                                $UcsCommCimcWebService = Get-UcsCommCimcWebService -Ucs $UCSM
                                if ($UcsCommCimcWebService) {
                                    Section -Style Heading4 -Name 'CIMC Web Service' {
                                        $UcsCimcWebService = [PSCustomObject] @{
                                            'State' = $UcsCommCimcWebService.AdminState
                                            'Port' = $UcsCommCimcWebService.Port
                                            'Operational Port' = $UcsCommCimcWebService.OperPort
                                            'Protocol' = $UcsCommCimcWebService.Proto
                                        }
                                        $UcsCimcWebService | Table -List -Name 'CIMC Web Service' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion CIMC Web Service Section

                                #region HTTP Section
                                $UcsHttp = Get-UcsHttp -Ucs $UCSM
                                if ($UcsHttp) {
                                    Section -Style Heading4 -Name 'HTTP' {
                                        $Http = [PSCustomObject]@{
                                            'State' = $UcsHttp.AdminState
                                            'Port' = $UcsHttp.Port
                                            'Operational Port' = $UcsHttp.OperPort
                                            'Request Timeout (in seconds)' = $UcsHttp.RequestTimeout
                                            'Redirect HTTP to HTTPS' = $UcsHttp.RedirectState
                                        }
                                        $Http | Table -List -Name 'HTTP Configuration' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion HTTP Section

                                #region HTTPS Section
                                $UcsHttps = Get-UcsHttps -Ucs $UCSM
                                if ($UcsHttps) {
                                    Section -Style Heading4 -Name 'HTTPS' {
                                        $Https = [PSCustomObject]@{
                                            'State' = $UcsHttps.AdminState
                                            'Port' = $UcsHttps.Port
                                            'Operational Port' = $UcsHttps.OperPort
                                            'Key Ring' = $UcsHttps.KeyRing
                                            'Cipher Suite Mode' = $UcsHttps.CipherSuiteMode
                                            'Cipher Suite' = $UcsHttps.CipherSuite
                                            'Allowed SSL Protocols' = $UcsHttps.AllowedSSLProtocols
                                        }
                                        $Https | Table -List -Name 'HTTPS Configuration' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion HTTPS Section

                                #region Telnet Section
                                $UcsTelnet = Get-UcsTelnet -Ucs $UCSM
                                if ($UcsTelnet) {
                                    Section -Style Heading4 -Name 'Telnet' {
                                        $Telnet = [PSCustomObject]@{
                                            'State' = $UcsTelnet.AdminState
                                            'Port' = $UcsTelnet.Port
                                            'Operational Port' = $UcsTelnet.OperPort
                                            'Protocol' = $UcsTelnet.Proto
                                        }
                                        $Telnet | Table -List -Name 'Telnet' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion Telnet Section

                                #region CIM XML Section
                                $UcsCimXml = Get-UcsCimXml -Ucs $UCSM
                                if ($UcsCimXml) {
                                    Section -Style Heading4 -Name 'CIM XML' {
                                        $CimXml = [PSCustomObject]@{
                                            'State' = $UcsCimXml.AdminState
                                            'Port' = $UcsCimXml.Port
                                            'Operational Port' = $UcsCimXml.OperPort
                                            'Protocol' = $UcsCimXml.Proto
                                        }
                                        $CimXml | Table -List -Name 'CIM XML' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion CIM XML Section

                                #region SNMP Section
                                $UcsSnmp = Get-UcsSnmp -Ucs $UCSM
                                if ($UcsSnmp) {
                                    Section -Style Heading4 -Name 'SNMP' {
                                        #TODO: Complete SNMP Config
                                        #region SNMP Configuration
                                        $Snmp = [PSCustomObject]@{
                                            'State' = $UcsSnmp.AdminState
                                            'Port' = $UcsSnmp.Port
                                            'Operational Port' = $UcsSnmp.OperPort
                                            'Protocol' = $UcsSnmp.Proto
                                            'System Contact' = $UcsSnmp.SysContact
                                            'System Location' = $UcsSnmp.SysLocation
                                        }
                                        $Snmp | Table -List -Name 'SNMP Configuration' -ColumnWidths 50, 50
                                        #endregion SNMP Configuration
                                        
                                        <#
                                        #region SNMP Traps
                                        #TODO: SNMP Traps
                                        $UcsSnmpTrap = Get-UcsSnmpTrap -Ucs $UCSM
                                        if ($UcsSnmpTrap) {
                                            Section -Style Heading4 -Name 'SNMP Traps' {
                                            }
                                        }
                                        #endregion SNMP Traps

                                        #region SNMP Users
                                        #TODO: SNMP Users
                                        $UcsSnmpUser = Get-UcsSnmpUser -Ucs $UCSM
                                        if ($UcsSnmpUser ) {
                                            Section -Style Heading4 -Name 'SNMP Users' {
                                            }
                                        }
                                        #endregion SNMP Users
                                        #>
                                    }
                                }
                                #endregion SNMP Section

                                #region SMASH CLP
                                $UcsSmashCLP = Get-UcsSmashCLP -Ucs $UCSM
                                if ($UcsSmashCLP) {
                                    Section -Style Heading4 -Name 'SMASH CLP' {
                                        $SmashCLP = [PSCustomObject]@{
                                            'State' = $UcsSmashCLP.AdminState
                                            'Port' = $UcsSmashCLP.Port
                                            'Operational Port' = $UcsSmashCLP.OperPort
                                            'Protocol' = $UcsSmashCLP.Proto
                                        }
                                        $SmashCLP | Table -List -Name 'SMASH CLP' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion SMASH CLP

                                #region SSH
                                $UcsSsh = Get-UcsSsh -Ucs $UCSM
                                if ($UcsSsh) {
                                    Section -Style Heading4 -Name 'SSH' {
                                        $Ssh = [PSCustomObject]@{
                                            'State' = $UcsSsh.AdminState
                                            'Port' = $UcsSsh.Port
                                            'Operational Port' = $UcsSsh.OperPort
                                            'Protocol' = $UcsSsh.Proto
                                        }
                                        $Ssh | Table -List -Name 'SSH' -ColumnWidths 50, 50
                                    }
                                }
                                #endregion SSH
                            }
                            #endregion Communication Services Section

                            #region DNS Management
                            $UcsDnsServer = Get-UcsDnsServer | Where-Object { ($_.Dn).StartsWith('org-root') }
                            if ($UcsDnsServer) {
                                Section -Style Heading4 -Name 'DNS Management' {
                                    $UcsDnsServer | Select-Object @{L = 'DNS Server'; E = { $_.Name } } | Table -Name 'DNS Management'
                                } 
                            }
                            #endregion DNS Management

                            #region Management Interfaces
                            $UcsMgmtInterfaceMonitorPolicy = Get-UcsMgmtInterfaceMonitorPolicy -Ucs $UCSM
                            if ($UcsMgmtInterfaceMonitorPolicy) {
                                Section -Style Heading4 -Name 'Management Interface Monitor Policy' {
                                    $UcsExtmgmtGatewayPing = Get-UcsExtmgmtGatewayPing -Ucs $UCSM
                                    $MgmtInterfaceMonitorPolicy = [PSCustomObject]@{
                                        'State' = $UcsMgmtInterfaceMonitorPolicy.AdminState
                                        'Poll Interval (seconds)' = $UcsMgmtInterfaceMonitorPolicy.PollInterval
                                        'Max Fail Report Count' = $UcsMgmtInterfaceMonitorPolicy.MaxFailReportCount
                                        'Monitor Mechanism' = $UcsMgmtInterfaceMonitorPolicy.MonitorMechanism
                                        'Number of Ping Requests' = $UcsExtmgmtGatewayPing.NumberOfPingRequests
                                        'Max Deadline Timeout (in seconds)' = $UcsExtmgmtGatewayPing.MaxDeadlineTimeout
                                    }
                                    $MgmtInterfaceMonitorPolicy | Table -List -Name 'Management Interface Monitor Policy' -ColumnWidths 50, 50
                                }
                            }
                            #endregion Management Interfaces

                            #region UCS Central
                            #TODO: Policy Resolution Control
                            $UcsCentral = Get-UcsCentral -Ucs $UCSM
                            if ($UcsCentral) {
                                Section -Style Heading4 -Name 'UCS Central' {
                                    $UcsCentralConfig = [PSCustomObject]@{
                                        'Hostname/IP Address' = $UcsCentral.SvcRegName
                                        'Repair State' = $UcsCentral.RepairState
                                        'Registration State' = $UcsCentral.RegistrationState
                                        'Cleanup Mode' = $UcsCentral.Cleanupmode
                                        'Suspend State' = Switch ($UcsCentral.SuspendState) {
                                            'off' { 'disabled' }
                                            'on' { 'enabled' }
                                        }
                                        'Acknowledge State' = Switch ($UcsCentral.AckState) {
                                            'no-ack' { 'disabled' }
                                            'ack' { 'enabled' }
                                        }
                                    }
                                    $UcsCentralConfig | Table -List -Name 'UCS Central' -ColumnWidths 50, 50
                                }
                            }
                            #endregion UCS Central

                        }
                        #endregion Communication Management Section

                        #region Stats Management
                        #TODO
                        #endregion Stats Management

                        #region Time Zone Management Section
                        $UcsTimeZone = Get-UcsTimezone -Ucs $UCSM | Select-Object -First 1
                        if ($UcsTimeZone) {
                            Section -Style Heading3 -Name 'Time Zone Management' {
                                $TimeZone = [PSCustomObject] @{
                                    'Time Zone' = Switch ($UcsTimeZone.TimeZone) {
                                        '' { 'not set' }
                                        default { $UcsTimeZone.TimeZone }
                                    }
                                }
                                $TimeZone | Table -List -Name 'Time Zone' -ColumnWidths 50, 50
                            }
                        }
                        #endregion Time Zone Management Section
                
                        #region License Management Section
                        $UcsLicenses = Get-UcsLicense -Ucs $UCSM
                        if ($UcsLicenses) {
                            Section -Style Heading3 -Name 'License Management' {
                                $Licenses = foreach ($UcsLicense in $UcsLicenses) {
                                    [PSCustomObject]@{
                                        'License Name' = $UcsLicense.Feature
                                        'Fabric' = $UcsLicense.Scope
                                        'Total Quantity' = $UcsLicense.AbsQuant
                                        'Used Quantity' = $UcsLicense.UsedQuant
                                        'Subordinate Used Quantity' = $UcsLicense.SubordinateUsedQuant
                                        'Default Quantity' = $UcsLicense.DefQuant
                                        'Operational State' = $UcsLicense.OperState
                                        'Grace Period Used' = "$($UcsLicense.GracePeriodUsed) days"
                                        'Peer License Count Comparison' = $UcsLicense.PeerStatus
                                    }
                                }
                                $Licenses | Sort-Object 'License Name', 'Fabric' | Table -Name 'License Management'
                            }
                        }
                        #region License Management Section
                    }
                }
                #endregion Admin Section
            }
            #endregion Script Body

            # Disconnect UCS Chassis
            $Null = Disconnect-Ucs -Ucs $UCSM
        }
    }
}