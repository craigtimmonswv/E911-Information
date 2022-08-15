<#
You will need have the "ImportExcel" Module installed for this to properly run. 
You can get it here:
https://www.powershellgallery.com/packages/ImportExcel/7.4.1
To install it run: 
Install-Module -Name ImportExcel -RequiredVersion 7.4.1
Import-Module -Name ImportExcel
This will pull the basic environment from the Teams tenant. Items it gathers is:
PSTN Gateways
PSTN Usages
Voice Routes
Voice Routing Policies
Dial Plan
Voice enabled users - this might take a while depending upon number of users
Emergency Calling Policies
Emergency Call Routing Policies
Tenant Network Site Details
LIS Locations
LIS Network Information
LIS WAP Information
LIS SWitch information
LIS Port
Auto Attendant
Call Queue
It will place the Excel spreadsheet it in the location you enter when prompted. 
#>

Function Write-DataToExcel
    {
        param ($filelocation, $details, $tabname)

        
        $excelpackage = Open-ExcelPackage -Path $filelocation 
        $ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName $tabname
        $details | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter
        Clv details 

    }
Function Get-LISDetails
{
            param ($filelocation)
            $Details = @()
            Write-Host "Running"
            
            # Get Emergency Calling Policies
            Write-Host 'Getting Emergency Calling Policies'
            $Details = @()
            $ercallpolicies = Get-CsTeamsEmergencyCallingPolicy
                foreach ($ercp in $ercallpolicies)
                {
                    $detail = New-Object PSObject
                    $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $ercp.Identity
                    $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $ercp.Description
                    $detail | add-Member -MemberType NoteProperty -Name "NotificationGroup" -Value $ercp.NotificationGroup
                    $detail | add-Member -MemberType NoteProperty -Name "ExternalLocationLookupMode" -Value $ercp.ExternalLocationLookupMode
                    $detail | add-Member -MemberType NoteProperty -Name "NotificationDialOutNumber" -Value $ercp.NotificationDialOutNumber
                    $detail | add-Member -MemberType NoteProperty -Name "NotificationMode" -Value $ercp.NotificationMode
                    $details += $detail  
                }
            $tabname = "Emergency Calling Policies"
            $Details|Export-Excel -Path $filelocation -WorksheetName "Emergency Calling Policies" -AutoSize -AutoFilter
            clv details
           # Write-DataToExcel $filelocation  $details $tabname

            # Get Emergency Call Routing Policy
            Write-Host 'Getting Emergency Call Routing Policies'
            $Details = @()
            $ecrps = Get-CsTeamsEmergencyCallRoutingPolicy
            foreach ($ecrp in $ecrps)
                {
                $numbers = Get-CsTeamsEmergencyCallRoutingPolicy -Identity $ecrp.identity
                foreach ($number in $numbers.EmergencyNumbers)
                    {
                        $detail = New-Object PSObject
                        $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $ecrp.Identity
                        $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $ecrp.Description
                        $detail | add-Member -MemberType NoteProperty -Name "emergencydialstring" -Value $number.emergencydialstring
                        $detail | add-Member -MemberType NoteProperty -Name "EmergencyDialMask" -Value $number.emergencydialmask
                        $detail | add-Member -MemberType NoteProperty -Name "OnlinePSTNUsage" -Value $number.OnlinePSTNUsage
                        $detail | add-Member -MemberType NoteProperty -Name "AllowEnhancedEmergencyServices" -Value $ecrp.AllowEnhancedEmergencyServices
                        $details  += $detail  
                    }
            }
            $tabname = "Emergency Call Routing Policies"
            Write-DataToExcel $filelocation  $details $tabname

            # Get Tenant Network Site Details
            Write-Host 'Getting Tenant Network Site Details'
            $Details = @()
            $erlocations = Get-CsTenantNetworkSite
            foreach ($location in $erlocations)
            {
                $networks = Get-CsTenantNetworkSubnet | ? {$_.networksiteid -eq $location.NetworkSiteID}
                foreach ($net in $networks)

                    {
                        $detail = New-Object PSObject
                        $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $location.Identity
                        $detail | add-Member -MemberType NoteProperty -Name "NetworkSiteID" -Value $net.NetworkSiteID
                        $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $net.Description
                        $detail | add-Member -MemberType NoteProperty -Name "SubnetID" -Value $net.SubnetID
                        $detail | add-Member -MemberType NoteProperty -Name "MaskBits" -Value $net.MaskBits
                        $detail | add-Member -MemberType NoteProperty -Name "EmergencyCallRoutingPolicy" -Value $location.EmergencyCallRoutingPolicy
                        $detail | add-Member -MemberType NoteProperty -Name "EmergencyCallingPolicy" -Value $location.EmergencyCallingPolicy
                        $details += $detail  
                    }
            }
            $tabname = "Tenant Network Site Details"
            Write-DataToExcel $filelocation  $details $tabname

            # Get Emergency Location information Services 
            Write-Host 'Getting Emergency Location Information Services'
            $locations = Get-CsOnlineLisLocation
            $Details = @()
            Foreach ($loc in $locations)
            {
                $detail = New-Object PSObject
                $detail | Add-Member NoteProperty -Name "CompanyName" -Value $loc.CompanyName
                $detail | Add-Member NoteProperty -Name "Civicaddressid" -Value $loc.civicaddressid
                $detail | Add-Member NoteProperty -Name "locationid" -Value $loc.LocationId
                $detail | Add-Member NoteProperty -Name "Description" -Value $loc.Description
                $detail | Add-Member NoteProperty -Name "location" -Value $loc.location
                $detail | Add-Member NoteProperty -Name "HouseNumber" -Value $loc.HouseNumber
                $detail | Add-Member NoteProperty -Name "HouseNumberSuffix" -Value $loc.HouseNumberSuffix
                $detail | Add-Member NoteProperty -Name "PreDirectional" -Value $loc.PreDirectional
                $detail | Add-Member NoteProperty -Name "StreetName" -Value $loc.StreetName
                $detail | Add-Member NoteProperty -Name "PostDirectional" -Value $loc.PostDirectional
                $detail | Add-Member NoteProperty -Name "StreetSuffix" -Value $loc.StreetSuffix
                $detail | Add-Member NoteProperty -Name "City" -Value $loc.City
                $detail | Add-Member NoteProperty -Name "StateOrProvince" -Value $loc.StateOrProvince
                $detail | Add-Member NoteProperty -Name "PostalCode" -Value $loc.PostalCode
                $detail | Add-Member NoteProperty -Name "Country" -Value $loc.CountryOrRegion
                $detail | Add-Member NoteProperty -Name "Latitude" -Value $loc.Latitude
                $detail | Add-Member NoteProperty -Name "Longitude" -Value $loc.Longitude
                $Details += $detail
                }
            $tabname = "LIS Location"
            Write-DataToExcel $filelocation  $details $tabname

            # Get LIS Network information
            Write-Host 'Getting LIS Network Information'
            $subnets = Get-CsOnlineLisSubnet
            $Details = @()
            Foreach ($subnet in $subnets)
            {
                $detail = New-Object PSObject
                $detail | Add-Member NoteProperty -Name "Subnet" -Value $subnet.Subnet
                $detail | Add-Member NoteProperty -Name "Description" -Value $subnet.Description
                $subloc = Get-CsOnlineLisLocation -LocationId $subnet.LocationId
                $detail | Add-Member NoteProperty -Name "Location" -Value $subloc.location
                $detail | Add-Member NoteProperty -Name "City" -Value $subloc.city
                $Details += $detail
            }
            $tabname = "LIS Network "
            Write-DataToExcel $filelocation  $details $tabname

            #Get LIS Wireless Access Point information
            Write-Host 'Getting LIS WAP Information'
            $WAPs = Get-CsOnlineLisWirelessAccessPoint
            $Details = @()
            Foreach ($WAP in $WAPs)
            {
                $detail = New-Object PSObject
                $detail | Add-Member NoteProperty -Name "BSSID" -Value $WAP.BSSID
                $detail | Add-Member NoteProperty -Name "Description" -Value $WAP.Description
                $WAPloc = Get-CsOnlineLisLocation -LocationId $WAP.LocationId
                $detail | Add-Member NoteProperty -Name "Location" -Value $WAPloc.location
                $detail | Add-Member NoteProperty -Name "City" -Value $WAPloc.city
                $Details += $detail
            }
            $tabname = "LIS WAP"
            Write-DataToExcel $filelocation  $details $tabname

            #Get LIS Switch information
            Write-Host 'Getting LIS SWitch information'
            $Switches = Get-CsOnlineLisSwitch
            $Details = @()
            Foreach ($Switch in $Switches)
            {
                $detail = New-Object PSObject
                $detail | Add-Member NoteProperty -Name "ChassisID" -Value $Switch.ChassisID
                $detail | Add-Member NoteProperty -Name "Description" -Value $Switch.Description
                $Switchloc = Get-CsOnlineLisLocation -LocationId $Switch.LocationId
                $detail | Add-Member NoteProperty -Name "Location" -Value $Switchloc.location
                $detail | Add-Member NoteProperty -Name "City" -Value $Switchloc.city
                $Details += $detail
                }
            $tabname = "LIS Switch"
            Write-DataToExcel $filelocation  $details $tabname

            #Get LIS Port information
            Write-Host 'Getting LIS Port Information'
            $Ports = Get-CsOnlineLisPort
            $Details = @()
            Foreach ($port in $ports)
                {
                $detail = New-Object PSObject
                $detail | Add-Member NoteProperty -Name "ChassisID" -Value $port.ChassisID
                $detail | Add-Member NoteProperty -Name "PortID" -Value $port.PortID
                $detail | Add-Member NoteProperty -Name "Description" -Value $port.Description
                $portloc = Get-CsOnlineLisLocation -LocationId $port.LocationId
                $detail | Add-Member NoteProperty -Name "Location" -Value $portloc.location
                $detail | Add-Member NoteProperty -Name "City" -Value $portloc.city
                $Details += $detail
            }
            $tabname = "LIS Port"
            Write-DataToExcel $filelocation  $details $tabname
            
            #Trusted IPs
            $trustedIPs = Get-CsTenantTrustedIPAddress
            $Details = @()
            Foreach ($IP in $TrustedIPs)
                {
                $detail = New-Object PSObject
                $detail | Add-Member NoteProperty -Name "Identity" -Value $IP.Identity
                $detail | Add-Member NoteProperty -Name "MaskBits" -Value $IP.MaskBits
                $detail | Add-Member NoteProperty -Name "Description" -Value $IP.Description
                $detail | Add-Member NoteProperty -Name "Description" -Value $IP.Description
                $Details += $detail
            }
            $tabname = "Trusted IPs"
            Write-DataToExcel $filelocation  $details $tabname

            
}
cls
Write-Host "This is will create an Excel Spreadsheet.  Make sure to enter the file name with .xlsx"
Import-Module ImportExcel
$filelocation = Read-Host "Enter Location/filename to store output (i.e c:\scripts\test.xlsx)"

# Determine if ImportExcel module is loaded


$XLmodule = Get-Module -Name importexcel



if ($XLmodule )
    {
        If ( $connected=get-cstenant -ErrorAction SilentlyContinue)
        {
            write-host "Current Tenant:" $connected.displayname
            Get-LISDetails $filelocation
        }
                Else {Write-Host "Teams module isn't loaded.  Please load Teams Module (connect-microsoftteams)"  }
    }
    Else { 
            Try {Import-Module ImportExcel -ErrorAction SilentlyContinue}
            Catch {Write-Host "ImportExcel module is not loaded"}
        }
