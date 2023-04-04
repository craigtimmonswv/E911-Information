# E911-Information
You will need have the "ImportExcel" Module installed for this to properly run. 
You can get it here:
https://www.powershellgallery.com/packages/ImportExcel/7.8.4
To install it run: 
Install-Module -Name ImportExcel -RequiredVersion 7.8.4
Import-Module -Name ImportExcel
This script will pull the basic environment from the Teams tenant. Items it gathers is:
Infrastructure items and makes the tabs green:
    Tenant Information
    PSTN Gateways
    PSTN Usages
    Voice Routes
    Voice Routing Policies
    Dial Plans
    Teams Meeting Settings.  If QOS is enabled, it will make the cell green.  If it isn't enabled, it will make the cell red. 
User details, and various policies.  It will make the tabs blue.  Items it gathers is:
    Voice enabled users - this might take a while depending upon number of users.  It is optional.
        Displayname, UPN, City, State, Country, Usage Location, Lineuri, Licenses, Dial Plan, Voice routing policy, Enterprise voice enabled,
        Teams upgrade policy, teams effective mode, emergency calling policy, emergency call routing policy, 
        Teams calling policy, Teams meeting policy, and Audio Conferencing Policy
    Auto-Attendant details
    Call Queue details
    Resource account details
    Caller ID Policy
    Calling Policies
    Audio Conferencing policies
Emergency services items
    Emergency Calling Policies
    Emergency Call Routing Policies
    Tenant Network Site Details
    LIS Locations
    LIS Subnets
    LIS Network Information
    LIS WAP Information
    LIS SWitch information
    LIS Port information

You will be prompted to enter a location to store the spreadsheet. This will be directory location something like "C:\scriptoutput".  It will then create a folder
called "TeamsEnvironmentReports".  This folder will hold the output of the spreadsheet and any error logs.  
The spreadsheet will have a name that contains the tenant, and date/time stamp.  It will look like "Contoso-TeamsEnv-11-18-2022.12.49.11.xlsx".
A few changes have been made:
    1. Format of the voice routing policy will have the OnlinePstnUages on one line, separted by commas.  This will allow reading of the sheet. 
    2. Added Error checking.  A log file will be created with the name similar to the file name, but will include errorlog in the filename.
    3. Added various policy reports (Meeting, Calling Policies, Caller ID Policy, Application Permission Policy, Teams Meeting Configuration, Audio Conferencing)
    4. Added feature types (licenses), assigned plans, and various Teams policies to the EV users report.  Some of these will be empty unless the user is using 
    calling plans, operator connect, DRaaS, or doing something with Video Interop.  
