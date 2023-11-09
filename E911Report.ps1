<#
This script will export LIS details from the the Microsoft Teams Tenant, 
create an excel spreadsheet.

Use Doug Finke ImportExcel Module located at https://github.com/dfinke/ImportExcel is required. 

You will be prompted for a directory location to store the report.  The report have a file name 
in the format of tenant name-E911-Date-Time-Standpoint.xlsx and will be stored in the directory
entered above.     

It will create the following tabs in the spreadsheet. 

    Tenant info
    LIS Location - Civic Address information and Location/Place information. 
    LIS Network (Subnet) information
    LIS WAP - Wireless Access Point BSSIDs
    LIS Switch 
    LIS Port - Switch Port information
    Trusted IP addresses
    Tenant Network Site Details
    Emergency Calling Policies
    Emergency Call Routing Policies

Additions

General
    -Added color coded tabs. 

BSSIDs
    - changed description to WAP-Description
    - Added Location Description

Added to tenant network subnet section
    - emergency calling policy with associated subnet
    - emergency call routing policy with associated subnet
    - subnet masks to tenant network subnet section
#>

Function Write-DataToExcel
    {
        param ($filelocation, $details, $tabname, $tabcolor)
        $excelpackage = Open-ExcelPackage -Path $filelocation 
        $ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName $tabname 
        $ws.Workbook.Worksheets[$ws.index].TabColor = $tabcolor
        $details | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter 
        Clear-Variable details 
        Clear-Variable filelocation
        Clear-Variable tabname
        Clear-Variable TabColor
    }
Function write-Errorlog
{
    param ($logfile, $errordata, $msgData)
    $errordetail = '"'+ $date + '","' + $msgData + '","' + $errordata + '"'
    Write-Host $errordetail
    $errordetail |  Out-File -FilePath $logname -Append 
    Clear-Variable errordetail, msgData
}
Function Write-TenantInfo
{
    param ($filelocation, $logfile)
    Write-host "Getting Tenant Information"
    $tenatDetail = Get-CsTenant
    $detail = New-Object PSObject
    $detail | add-Member -MemberType NoteProperty -Name "DisplayName" -Value $tenatDetail.DisplayName
    $detail | add-Member -MemberType NoteProperty -Name "TeamsUpgradeEffectiveMode" -Value $tenatDetail.TeamsUpgradeEffectiveMode
    $detail | add-Member -MemberType NoteProperty -Name "TenantId" -Value $tenatDetail.TenantId
    $Detail |Export-Excel -Path $filelocation -WorksheetName "Tenant info" -AutoFilter -AutoSize
    $excel = Open-ExcelPackage -Path $filelocation 
    $Green = "Green"
    $Green = [System.Drawing.Color]::$green 
    $excel.Workbook.Worksheets[1].TabColor = $Green  
    Close-ExcelPackage -ExcelPackage $excel
}

Function Write-EmergencyCallingPolicy
{
    param ($filelocation, $logfile)
    # Get Emergency Calling Policies
    
    Write-Host 'Getting Emergency Calling Policies'
    $Details = @()
    try {$ercallpolicies = Get-CsTeamsEmergencyCallingPolicy -ErrorAction Stop }
    catch 
    {
        $msgdata = "Error getting Emergency Calling Policy Details."
        write-Errorlog $logfile $error[0].exception.message $msgData
        Clear-Variable msgData
    }
    if ($ercallpolicies -ne 0)
    {
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
    }
    Else {$details = "No Data to Display"}
    $tabname = "Emergency Calling Policies"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-EmergencyCallRouting
{
    param ($filelocation, $logfile)
    # Get Emergency Call Routing Policy
    Write-Host 'Getting Emergency Call Routing Policies'
    $Details = @()
    try {$ecrps = Get-CsTeamsEmergencyCallRoutingPolicy -ErrorAction Stop }
    catch 
    {
        $msgdata = "Error getting Emergency Call Routing Policy Details."
        write-Errorlog $logfile $error[0].exception.message $msgData
        Clear-Variable msgData
    }
    if ($ecrps.count -ne 0)
    {
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
    }
    Else {$details = "No Data to Display"}
    $tabname = "Emergency Call Routing Policies"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-NetworkSiteDetails
{
    param ($filelocation, $logfile)
    # Get Tenant Network Site Details
    Write-Host 'Getting Tenant Network Site Details'
    $Details = @()
    try {$sites = Get-CsTenantNetworkSite -ErrorAction Stop}
    catch 
        {
            $msgdata = "Error getting Tenant Network Site Details."
            write-Errorlog $logfile $error[0].exception.message $msgData
            Clear-Variable msgData
        }
    if ($sites.count -ge 1)
        {
            foreach ($site in $sites)
            {
                $detail = New-Object PSObject
                $detail | add-Member -MemberType NoteProperty -Name "Subnets" -Value $site.Subnets
                $detail | add-Member -MemberType NoteProperty -Name "Postalcodes" -Value $site.Postalcodes
                $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $site.Identity
                $detail | add-Member -MemberType NoteProperty -Name "NetworkSiteID" -Value $site.NetworkSiteID
                $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $site.Description
                $detail | add-Member -MemberType NoteProperty -Name "NetworkRegionID" -Value $site.NetworkRegionID
                $detail | add-Member -MemberType NoteProperty -Name "LocationPolicy" -Value $site.LocationPolicy
                $detail | add-Member -MemberType NoteProperty -Name "EnableLocationBasedRouting" -Value $site.EnableLocationBasedRouting
                $detail | add-Member -MemberType NoteProperty -Name "SiteAddress" -Value $site.SiteAddress
                $detail | add-Member -MemberType NoteProperty -Name "EmergencyCallRoutingPolicy" -Value $site.EmergencyCallRoutingPolicy
                $detail | add-Member -MemberType NoteProperty -Name "EmergencyCallingPolicy" -Value $site.EmergencyCallingPolicy
                $detail | add-Member -MemberType NoteProperty -Name "NetworkRoamingPolicy" -Value $site.NetworkRoamingPolicy
                $details += $detail  
            }
        }
    
    Else {$details = "No Data to Display"}
    $tabname = "Tenant Network Site Details"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}
Function Write-NetworkRegion
{  
    param ($filelocation, $logfile) 
    Write-Host "Getting Tenant Network Region"
    $Details = @()
    $regions = Get-CsTenantNetworkRegion
    if ($regions.count -ge 1)
    {
        foreach ($region in $regions)
        {
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $region.Identity
            $detail | add-Member -MemberType NoteProperty -Name "NetworkRegionID" -Value $region.NetworkRegionID
            $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $region.Description
            $detail | add-Member -MemberType NoteProperty -Name "CentralSite" -Value $region.CentralSite
            $Details += $detail
        }
    }
    else {$details = "No Data to Display"}
    $tabname = "Tenant Network Region"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}


Function Write-NetworkSubnetDetails
{ 
    param ($filelocation, $logfile)  
    Write-Host "Getting Tenant Network Subnets"
    $Details = @()
    $subnets = Get-CsTenantNetworkSubnet
    if ($subnets.count -ge 1)
    {
        foreach ($subnet in $subnets)
        {
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $subnet.Description
            $detail | add-Member -MemberType NoteProperty -Name "Subnet" -Value $subnet.SubnetID
            $detail | add-Member -MemberType NoteProperty -Name "Masks" -Value $subnet.MaskBits
            $detail | add-Member -MemberType NoteProperty -Name "NetworkSiteID" -Value $subnet.NetworkSiteID
            $site = Get-CsTenantNetworkSite -Identity $subnet.NetworkSiteID
            try {
                $ECP = Get-CsTeamsEmergencyCallingPolicy -Identity $site.EmergencyCallingPolicy -ErrorAction Stop
                $ECPIdentity = $ecp.Identity
                $ecpXtrnLocationLookupMode = $ecp.ExternalLocationLookupMode
                $ecpNotificationMode = $ecp.NotificationMode
                $ecpnotificationgroup = $ecp.NotificationGroup
                $ecpNotificationDialOutNumber = $ecp.NotificationDialOutNumber
            
            }
            catch {}
            if (!($ecp))
            {
                $ECPIdentity = "Null"
                $ecpXtrnLocationLookupMode = "Null"
                $ecpNotificationMode = "Null"
                $ecpnotificationgroup = "Null"
                $ecpNotificationDialOutNumber = "Null"

        }
            $detail | add-Member -MemberType NoteProperty -Name "ECP Identity" -Value $ECPIdentity
            $detail | add-Member -MemberType NoteProperty -Name "External Location Lookup" -Value $ecpXtrnLocationLookupMode
            $detail | add-Member -MemberType NoteProperty -Name "Notification Mode" -Value $ecpNotificationMode
            $detail | add-Member -MemberType NoteProperty -Name "Notification Group" -Value $ecpnotificationgroup
            $detail | add-Member -MemberType NoteProperty -Name "Notification Dial OutNumber" -Value $ecpNotificationDialOutNumber
            $Details += $detail
        }
    }
    else {$details = "No Data to Display"}
    $tabname = "Tenant Network Subnet"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-TrustedIPs
{
    param ($filelocation, $logfile)
    # Get Tenant Trusted IP Addresses
    Write-Host 'Getting Tenant Trusted IP Addresses'
    $Details = @()
    try {$TrustedIPs = get-CsTenantTrustedIPAddress -ErrorAction Stop}
    catch 
        {
            $msgdata = "Error getting Trusted IP Address Details."
            write-Errorlog $logfile $error[0].exception.message $msgData
            Clear-Variable msgData
        }
    if ($TrustedIPs.count -ne 0)
    {
        foreach ($TrustedIP in $TrustedIPs)
        {
            $IP = get-CsTenantTrustedIPAddress | Where-Object {$_.IPAddress -eq $TrustedIP.IPAddress}
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $IP.Identity
            $detail | add-Member -MemberType NoteProperty -Name "IPAddress" -Value $IP.IPAddress
            $detail | add-Member -MemberType NoteProperty -Name "MaskBits" -Value $IP.MaskBits
            $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $IP.Description
            $details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "Trusted IP address"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-LISLocation
{
    param ($filelocation, $logfile)
    # Get Emergency Location information Services 
    Write-Host 'Getting Emergency Location Information Services'
    $locations = Get-CsOnlineLisLocation
    if ($locations.count -ne 0)
    {
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
    }
    Else {$details = "No Data to Display"}
    $tabname = "LIS Location"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-LISSubnets
{
    param ($filelocation, $logfile)
    # Get LIS Network information
    Write-Host 'Getting LIS Network Information'
    try {$subnets = Get-CsOnlineLisSubnet -erroraction Stop}
    catch 
        {   
            $msgdata = "Error getting LIS Subnets Details."
            write-Errorlog $logfile $error[0].exception.message $msgData
            Clear-Variable msgData
        }
    if ($subnets.count -ne 0)
        {
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
    }
    Else {$details = "No Data to Display"}
    $tabname = "LIS Network"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-BSSIDs
{
    param ($filelocation, $logfile)
    #Get LIS Wireless Access Point information
    Write-Host 'Getting LIS WAP Information'
    try {$WAPs = Get-CsOnlineLisWirelessAccessPoint -ErrorAction Stop}
    catch 
    {
        $msgdata = "Error getting LIS WAP Details."
        write-Errorlog $logfile $error[0].exception.message $msgData
        Clear-Variable msgData
    }
    if ($waps.count -ne 0)
    {
        $Details = @()
        Foreach ($WAP in $WAPs)
        {
            $detail = New-Object PSObject
            $detail | Add-Member NoteProperty -Name "BSSID" -Value $WAP.BSSID
            $detail | Add-Member NoteProperty -Name "WAP-Description" -Value $WAP.Description
            $WAPloc = Get-CsOnlineLisLocation -LocationId $WAP.LocationId
            $detail | Add-Member NoteProperty -Name "Location" -Value $WAPloc.location
            $detail | Add-Member NoteProperty -Name "Location Description" -Value $WAPloc.description
            $detail | Add-Member NoteProperty -Name "City" -Value $WAPloc.city
            $Details += $detail
        }
    }
    Else {$details = "No Data to Display"}
    $tabname = "LIS WAP"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-LISSwitch
{
    param ($filelocation, $logfile)
    #Get LIS Switch information
    Write-Host 'Getting LIS SWitch information'
    $Switches = Get-CsOnlineLisSwitch -ErrorAction Stop
    $Details = @()
    if ($Switches.count -ne 0)
    {
        
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
    }
    Else {$details = "No Data to Display"}
    $tabname = "LIS Switch"
    $tabcolor = "Red"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Function Write-LISPort
{
    param ($filelocation, $logfile)
     #Get LIS Port information
     Write-Host 'Getting LIS Port Information'
     try {$Ports = Get-CsOnlineLisPort -ErrorAction stop}
     catch 
     {
         $msgdata = "Error getting LIS Port Details."
         write-Errorlog $logfile $error[0].exception.message $msgData
         Clear-Variable msgData
     }
     if ($ports.count -gt 0)
     {
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
     }
     else {$details = "No data to display"}
     $tabname = "LIS Port"
     $tabcolor = "Red"
     Write-DataToExcel $filelocation $Details $tabname $tabcolor
}

Import-Module ImportExcel

# Determine if ImportExcel module is loaded
Clear-Host
$XLmodule = Get-Module -Name importexcel
if ($XLmodule )
    {
        If ( $connected=get-cstenant -ErrorAction SilentlyContinue)
            {
                Write-Host "This is will create an Excel Spreadsheet."
                $dirlocation = Read-Host "Enter location to store report (i.e. c:\scriptout)"
                Clear-Host
                $directory = $dirlocation+"\E911"
                try { Resolve-Path -Path $directory -ErrorAction Stop }
                catch 
                    {
                        Try {new-item -path $directory -itemtype "Directory" -ErrorAction Stop}
                        Catch 
                        {
                            $logfile, $errordata, $msgData
                            $date = get-date -Format "MM/dd/yyyy HH:mm"
                            $errordetail = $date + ", Error creating directory. ," + $directory+ ","+ $error[0].exception.message 
                            Write-Host $errordetail
                        }
                    }
                
                write-host "Current Tenant:" $connected.displayname
                $filedate=Get-Date -Format "MM-dd-yyyy.HH.mm.ss"
                $tenant = $connected.displayname.Replace(" ","-")
                $filelocation = $directory+"\"+$tenant+"-E911-"+$filedate+".xlsx"
                $logfile = $directory+"\"+$tenant+"-TeamsEnv-ErrorLog-"+$filedate+".csv"
                Write-TenantInfo $filelocation $logfile
                Write-EmergencyCallingPolicy $filelocation $logfile
                Write-EmergencyCallRouting $filelocation $logfile
                Write-NetworkRegion $filelocation $logfile
                Write-NetworkSiteDetails $filelocation $logfile
                Write-NetworkSubnetDetails $filelocation $logfile
                Write-TrustedIPs $filelocation $logfile
                Write-LISLocation $filelocation $logfile
                Write-LISSubnets $filelocation $logfile
                Write-BSSIDs $filelocation $logfile
                Write-LISSwitch $filelocation $logfile
                Write-LISPort $filelocation $logfile
            }
        Else {Write-Host "Teams module isn't loaded.  Please load Teams Module (connect-microsoftteams)"  }
    }
Else {Write-Host "ImportExcel module is not loaded"}