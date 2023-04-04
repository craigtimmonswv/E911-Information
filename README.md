# E911-Information

This script will export LIS details from the the Microsoft Teams Tenant, 
create an excel spreadsheet.

Use of Doug Finke ImportExcel Module located at https://github.com/dfinke/ImportExcel is required. 

You will be prompted for a directory location to store the report.  The report have a file name 
in the format of tenant name-E911-Date-Time-Stamp.xlsx (Contoso-E911-04-04-2023.16.35.13.xlsx) 
and will be stored in the directory entered above.     

It will create the following tabs in the spreadsheet. <br>
    - Tenant info<br>
    - LIS Location - Civic Address information and Location/Place information. <br>
    - LIS Network (Subnet) information<br>
    - LIS WAP - Wireless Access Point BSSIDs<br>
    - LIS Switch <br>
    - LIS Port - Switch Port information<br>
    - Trusted IP addresses<br>
    - Tenant Network Site Details<br>
    - Emergency Calling Policies<br>
    - Emergency Call Routing Policies<br>
