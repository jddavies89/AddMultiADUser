
<#

.SYNOPSIS
    This script displays Office365 user data from your tenancy.

.DESCRIPTION    
    This getmeTenancy script is an advanced module, please see 'PARAMETER's and'EXAMPLES' on what features are available and how to use the script.


    Functions to edit before running the script for the first time:
        addSellectedProPlusFacultyUsr; Here 'Office 365 ProPlus for faculty' get enabled with only 'SWAY' license disabled.
        addSellectedProPlusStudentsUsr; Here 'Office 365 ProPlus for Students' gets enabled with only 'SWAY' and 'DFFICE_FORMS_PLAN_2' lcenses disabled.
        addSellectedEduFacultyUser; Here 'Office 365 Education for faculty' gets enabled with School Data Sync (plan1), Stream for Office365, Microsoft Teams, PowerApps for Office365, Azure Rights Management licenses disabled.
        addSellectedEduStudentUsr; Here 'Office365 for student education' gets enabled with Microsoft Forms (plan 2), Microsoft Planner, SWAY, Office Online for Education, Sharepoint 1 for EDU, Exchange Online (plan 1)


    For a list of parameters in runtime;
        get-help GetMeTenancyDisplay -Parameter *
        get-help GetMeTenancyWrite -Parameter *


    Modules:
        MSONline

.PARAMETER -Connect
    Connects to Office365 and prompts for the password, this has to be with -O365User.
.PARAMETER -O365User jon.doe@office365tenancy.com
    Stores the username of Office365 logon tenancy.
.PARAMETER -O354Password Password_Of_Account
    Stores the Password for OFfice365 logon tenancy.
.PARAMETER -connectToO365Spec
    Connects to Office365 tenancy with specific username and password from -O365User and -O365Password.

.PARAMETER -searchUsr jon.doe@office365tenancy.com
    Searches for all users that your search term specifies.
.PARAMETER -displaySiteSKULicServicePlan
    Display the site SKU and Service plan for each SKU.
.PARAMETER -displayUsrLicense jon.doe@office365tenancy.com
    Displays the searched user's SKU and Service plans.
.PARAMETER -listAllUnlicensedUsrs
    List all the user's who have no licenses.
.PARAMETER -exportAllUnlicensedUsrs
    Export all the users who have no licenses with a specified path.

.PARAMETER -updateAllProPlusFacultyUsr jon.doe@office365tenancy.com
    Updates Office365 ProPlus for faculty, all licenses enabled for a specific user.
.PARAMETER -updateAllProPlusStudentUsr jon.doe@office365tenancy.com
    Updates Office365 ProPlus for Students, all licenses enabled for a specific user.
.PARAMETER -updateAllEduFacultyUsr jon.doe@office365tenancy.com
    Updates Office365 for faculty education, all licenses enabled for a specific user.
.PARAMETER -updateAllEduStudentUsr jon.doe@office365tenancy.com
    Updates Office365 for student education, all licenses enabled for a specific user.

.PARAMETER -addAllProPlusFacultyUsr jon.doe@office365tenancy.com
    Add Office365 ProPlus for faculty, all licenses enabled for a specific user.
.PARAMETER -addAllProPlusStudentsUsr jon.doe@office365tenancy.com
    Add Office365 ProPlus for Students, all licenses enabled for a specific user.
.PARAMETER -addAllEduFacultyUsr jon.doe@office365tenancy.com
    Add Office365 for faculty education, all licenses enabled for a specific user.
.PARAMETER -addAllEduStudentUsr jon.doe@office365tenancy.com
    Add Office365 for student education, all licenses enabled for a specific user.

.PARAMETER -addSellectedProPlusFacultyUsr
    Add Selected licenses under the SKU of Office365 ProPlus for faculty.
.PARAMETER -addSellectedProPlusStudentsUsr
    Add selected licenses under the SKU of Office365 ProPlus for students.
.PARAMETER addSellectedEduFacultyUsr
    Add sellected licenses under the Office365 for faculty education.
.PARAMETER addSellectedEduStudentUsr
    Add sellected licenses under the Office365 for student education.

.PARAMETER -importCSVUsersAndEditLicenses C:\test.cs
    Imports the csv file, this csv file has to have the EMail Address of each user on the tenancy. 
    This parameter has to be implemented with one of the following parameters and string of csv:

    .PARAMETER -updateAllProPlusFacultyUsr csv
    .PARAMETER -updateAllProPlusStudentUsr csv
    .PARAMETER -updateAllEduFacultyUsr csv
    .PARAMETER -updateAllEduStudentUsr csv

    .PARAMETER -addAllProPlusFacultyUsr csv
    .PARAMETER -addAllProPlusStudentsUsr csv
    .PARAMETER -addAllEduFacultyUsr csv
    .PARAMETER -addAllEduStudentUsr csv

    .PARAMETER -addSellectedProPlusFacultyUsr csv
    .PARAMETER -addSellectedProPlusStudentsUsr csv
    .PARAMETER addSellectedEduFacultyUsr csv
    .PARAMETER addSellectedEduStudentUsr csv

.EXAMPLE
    GetMeTenancyDisplay -Connect -O365Username "jon.doe@office365tenancy.com"
        Connects to Office365 with the username of jon.doe@office365tenancy.com and prompts for a password.
.EXAMPLE
    GetMeTenancyDisplay -connectToO365Spec -O365User "jon.doe@office365tenancy.com" -O365Password "Password_Of_Account"
        Connects to Office£65 with specific Username nad password credentials.
.EXAMPLE
    GetMeTenancyDisplay -connectToO365Spec -O365User "jon.doe@office365tenancy.com" -O365Password "Password_Of_Account" -searchUsr "jrichards" -disconnect
       Connects to OFfice365 with specific Credentials, searches for jrichards then Removes all PSSessions with *.outlook.com.
.EXAMPLE
    GetMeTenancyDisplay -connectToO365Spec "jon.doe@office365tenancy.com" -O365Password "Password_Of_Account" -displaySiteSKULicServicePlan -disconnect
        Connects to Office365 with specific username and password, displays the SKU license and service plan then disconnect.
.EXAMPLE
    GetMeTenancyDisplay -Connect -O365User "jon.doe@office365tenancy.com" -displaySiteSKULicServicePlan
        Connects to Office365 with a prompted password then displays the SKU license and service plans.
.EXAMPLE
    GetMeTenancyDisplay -O365User -connectToO365Spec "jon.doe@office365tenancy.com" -O365Password "Password_Of_Account" -connectToO365Spec -displayUsrLicense "jon.doe@office365tenancy.com"
        Connects to Office365 with no prompt and then display the user licenses for jon.doe@office365tenancy.com.
.EXAMPLE
    GetMeTenancyDisplay -exportAllUnlicensedUsrs "C:\unsignedUserLicenses.csv"
        Export all unsigned user's license to C:\unsignedUserLicenses.csv, has to be already connected to Office365.
.EXAMPLE
    GetMeTenancyDisplay -searchUsr "jrichards"
       Searches for jrichards on Office365.
.EXAMPLE
    GetMeTenancyWrite -updateO365AllLicenseToSingleUser jon.doe@office365tenancy.com
        Updates Office365 for faculty and enables all licenses for jon.doe@office365tenancy.com.
 .EXAMPLE
    GetMeTenancyWrite -connectToO365Spec -O365User "jon.doe@office365tenancy.com" -O365Password "Password_Of_Account" addAllProPlusFacultyUsr jon.doe@office365tenancy.com -disconnect
        Connects to Office365 with specific credentials, add all the licenses from ProPlus for faculty to jon.doe@office365tenancy.com then disconnects from the tenancy.
.EXAMPLE
    GetMeTenancyWrite -connectToO365Spec -O365User jon.doe@office365tenancy.com -O365Password "Password_Of_Account" -importCSVUsersAndEditLicenses C:\test.csv -updateAllProPlusFacultyUser csv
        Connects to Office365 with specified written credentials then updates all the ProPlus for Faculty licenses for each user imported from C:\test.csv.
.EXAMPLE
    GetMeTenancyWrite -importCSVUsersAndEditLicenses C:\test.csv -addSellectedEduFacultyUser csv
        Updates the sellected Office365 Education for faculty licenses for each user imported from C:\test.csv.

.Notes
   Author: Joe Richards
   Date:   15/09/2017
.LINK
  https://github.com/joer89/Admin-Tools.git
#>


#Reads data from Office365.
function GetMeTenancyDisplay {

    [CmdletBinding(SupportsShouldProcess=$true)]        


    param (
        [Parameter(Mandatory=$false)]
        [switch]$Connect,

        [parameter(Mandatory=$false)]
        [switch]$connectToO365Spec,
        
        [parameter(Mandatory=$false)]
        [switch]$disconnect,
       
        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$O365User,
        
        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$O365Password,
        
        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$searchUsr,

        [parameter(Mandatory=$false)]
        [switch]$displaySiteSKULicServicePlan,         
        
        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$displayUsrLicense,

        [parameter(Mandatory=$false)]
        [switch]$listAllUnlicensedUsrs,

        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$exportAllUnlicensedUsrs

        )   
        

        Begin{
            checkMSOnlinConnection
        }#End Begin
        process{
            #Searches for a specific user.
            if($searchUsr){                
                searchUser($searchUsr)
            }#End if
            #Display the sites SKU and Service licenses.
            if($displaySiteSKULicServicePlan){               
               displaySiteSKULicensesSerPlan
            }#End if
            #Display a specific user's licenses.
            if($displayUsrLicense){                
                displayUserLicense($displayUsrLicense)
            }#End if
            #Display all unsigned licenses.
            if($listAllUnlicensedUsrs){
                listAllUnlicensedUsers
            }#End if
            #Export to a CSV file of all unsigned user licenses and specifiy the CSV file path.
            if($exportAllUnlicensedUsrs){
                exportAllUnlicensedUsers($exportAllUnlicensedUsrs)
            }#End if
        }#End Process
        end{
            if($disconnect){
                #Disconnect all Outlook PSSessions.
                disconnect
            }#End if 
        }#End of end
    }

#Writes data to Office365.
function GetMeTenancyWrite {
 
  [CmdletBinding(SupportsShouldProcess=$true)]        


    param (
        [parameter(Mandatory=$false)]
        [switch]$connectToO365Spec,
        
        [parameter(Mandatory=$false)]
        [switch]$disconnect,
       
        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$O365User,
        
        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$O365Password,


        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$updateAllProPlusFacultyUser,

        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$updateAllProPlusStudentsUser,

        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$updateAllEduFacultyUser,
        
        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$updateAllEduStudentUser,

        
        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$addAllProPlusFacultyUser,

        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$addAllProPlusStudentsUser,

        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$addAllEduFacultyUser,
        
        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$addAllEduStudentUser,


        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$addSellectedProPlusFacultyUser,

        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$addSellectedProPlusStudentsUser,

        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$addSellectedEdcuFacultyUser,

        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$addSellectedEdcuStudentUser,


        [parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$importCSVUsersAndEditLicenses



        )
        begin{
            checkMSOnlinConnection
        }#End Begin)
        process{
            #Asigns all licenses under the SKU of Office365 ProPlus for faculty.
            if($updateAllProPlusFacultyUser){
               updateAllProPlusFacultyUsr($updateAllProPlusFacultyUser)
            }
            #Asigns all licenses under the SKU of Office365 ProPlus for students.
            if($updateAllProPlusStudentsUser){
               updateAllProPlusStudentUsr($updateAllProPlusStudentsUser)
            }
            #Assign all licenses under the SKU of Office 365 Education for faculty.
            if($updateAllEduFacultyUser){
                updateAllEduFacultyUsr($updateAllEduFacultyUser)
            }
            #Assign all licenses under the SKU of Office 365 Education for students.
            if($updateAllEduStudentUser){
                 updateAllEduStudentUsr($updateAllEduStudentUser)
            }

            #Add all licenses under the SKU of Office365 ProPlus for faculty.
            if($addAllProPlusFacultyUser){
               addAllProPlusFacultyUsr($addAllProPlusFacultyUser)
            }
            #Add all licenses under the SKU of Office365 ProPlus for students.
            if($addAllProPlusStudentsUser){
               addAllProPlusStudentsUsr($addAllProPlusStudentsUser)
            }
            #Add all licenses under the SKU of Office 365 Education for faculty.
            if($addAllEduFacultyUser){
               addAllEduFacultyUsr($addAllEduFacultyUser)
            }
            #Add all licenses under the SKU of Office 365 Education for students.
            if($addAllEduStudentUser){
               addAllEduStudentUsr($addAllEduStudentUser)
            }

             #Add sellected licenses under the SKU of Office365 ProPlus for faculty.
            if($addSellectedProPlusFacultyUser){
                addSellectedProPlusFacultyUsr($addSellectedProPlusFacultyUser)
            }
            #Add sellected licenses under the SKU of Office365 ProPlus for students.
            if($addSellectedProPlusStudentsUser){
               addSellectedProPlusStudentsUsr($addSellectedProPlusStudentsUser)
            }
            #Add sellected licenses under the SKU of Office 365 Education for faculty.
            if($addSellectedEduFacultyUser){
               addSellectedEduFacultyUsr($addSellectedEduFacultyUser)
            }  
            #Add sellected licenses under the SKU of Office 365 Education for students.
            if($addSellectedEdcuStudentUser){
                addSellectedEdcuStudentUsr($addSellectedEdcuStudentUser)
            }

            #Imports the email address of the users from a csv file to edit their licenses.
            if($importCSVUsersAndEditLicenses){
                importCSVUsrsAndEditLicenses($importCSVUsersAndEditLicenses)
             }
        }#End process
        end{
            if($disconnect){
                #Disconnect all Outlook PSSessions.
                disconnect
            }#End if 
        }#End of end
}


#Checks to see if the connection to Office365 is open.
function checkMSOnlinConnection{
#Checks to see if the MSOnline module is loaded.
           $MSOnlineMod = Get-Module -Name MSOnline | Format-Table Name -HideTableHeaders
           if($MSOnlineMod -ne "MSOnline"){            
                impModules
                if($Connect){
                    if($O365User -ne $null -or $O365User -ne ""){ 
                        #Connects to Office365 with manual credentials.                      
                        connectToO365($O365User)
                    }
                }
                elseif($connectToO365Spec){
                    if($O365User -ne $null -or $O365User -ne "" -and $O365Password -ne $null -or $O365Password -ne ""){
                        #Connects to Office365 with user's input Username and Password with no prompt.
                        connectToO365SpecifyUsrPwd($O365User)($O365Password)
                    }#End if
                }#End if
            }#End if
}
#Imports the modules.
function impModules{
    #Imports the MSOnline module for Office365.
    Import-Module MsOnline -ErrorAction Stop
     Write-Host "Imported MSOnline" -ForegroundColor Magenta
}
#Connects to Office365 and ask for credentials.
function connectToO365($O365User){
    try{
        $cred = Get-Credential $O365User
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
        Import-PSSession $session -AllowClobber
        Connect-MsolService -Credential $cred
        Write-Host "Connected to Office365" -ForegroundColor Magenta
    }
    catch{
        Write-Host "Connection failed."
    }
}
#Connect to Office365 with specific usernmae and password.
function connectToO365SpecifyUsrPwd($O365User,$O365Password){
    #Specified Credentials.
    #This is so you don't have to enter the password each time.
        try{
            $pass = ConvertTo-SecureString -AsPlainText $O365Password -Force
            $Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $O365User,$pass
       
            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Cred -Authentication Basic -AllowRedirection
            Import-PSSession $session -AllowClobber
            Connect-MsolService -Credential $Cred
            Write-Host "Connected to Office365"
    }
    catch{
            Write-Host "Connection failed."
    }
}
#Disconnects all PSSessions with *.outlook.com.
function disconnect{
    #Disconnects the PSSEssions with *.outlook
    Get-PSSession | ? {$_.ComputerName -like ".outlook.com"} | Remove-PSSession
    Write-Host "Removed all PSSessions with *.outlook.com" -ForegroundColor Magenta
}

#Searches for users on Office365.
function searchUser($user){
    Write-Host "Searching for $($user)" -ForegroundColor Magenta
    #Searches for a user in Office365 and display as a table view.
    Get-MsolUser -SearchString $user | Format-Table
    Write-Host "Finished searching for $($user)" -ForegroundColor Magenta
}
#Display the SKU and service plans for the SKU.
function displaySiteSKULicensesSerPlan{
    Write-Host "Displaying the sites SKU and Service plans." -ForegroundColor Magenta
    #Stores the SKUs for the company.
    $AccountSku = Get-MsolAccountSku
    #Display each license product under the SKUs.
    for ($i=0; $i -lt  $AccountSku.Count; $i++){
            $AccountSku[$i] | Format-Table
            $AccountSku[$i].ServiceStatus | Format-Table
    } 
    Write-Host "Finished displaying the sites SKU and Service plans." -ForegroundColor Magenta
}
#Displays the user's license and SKU.
function displayUserLicense($user){
    Write-Host "Displaying user's licenses" -ForegroundColor Magenta
    #Stores the SKUs for the company.
    $AccountSku = Get-MsolAccountSku

    #obtain userlicense.
    $userLicense = Get-MsolUser -UserPrincipalName $user
    #Return True or false.
    if($userLicense.isLicensed -eq "True"){
        Write-Host "$($user) is licensed."
    }
    elseif ($userLicense.IsLicensed -eq "False"){
        Write-Host "$($user) is not licensed."
    }

    #Displays the amount of licenses.
    Write-host "$($user) has $($userLicense.Licenses.Count) licenses."

    #Checks to see if there is any licenses.
    if($userLicense.Licenses.Count -gt 0){
        #Display a detailed list of product licenses under their SKU.
        for ($i=0; $i -lt $AccountSKU.Count; $i++){
            #Display the SKU.
            $AccountSku[$i] | Format-Table
            #Display the Licenses within the SKU for the user.
            $userLicense.Licenses[$i].ServiceStatus
        }
    }
     Write-Host "Finished displaying user's licenses" -ForegroundColor Magenta
}
#Display all users that are not licensed.
function listAllUnlicensedUsers{        
   try{
        Write-Host "Displaying all users who have no licenses." -ForegroundColor Magenta
        Get-MsolUser -All | where {$_.isLicensed -eq $false}| Format-list UserPrincipalName        
        Write-Host "Finished displaying all users who have no licenses." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to display unlicensed users to." -ForegroundColor Red
   }
}
#Display all users that are not licensed and export to a CSV file.
function exportAllUnlicensedUsers($exportPath){
    try{
        Write-Host "Exporting unlicensed users to $($exportPath)." -ForegroundColor Magenta
        Get-MsolUser -All | where {$_.IsLicensed -eq $false} | Select-Object UserPrincipalName | Export-Csv $exportPath
        Write-Host "Exported unlicensed users to $($exportPath)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to Export unlicensed users to $($exportPath)." -ForegroundColor Red
   }
}

#Updates Office365 ProPlus for faculty, all licenses enabled for a specific user.
function updateAllProPlusFacultyUsr($user){
   try{
        Write-Host "Enabling all of 'Office 365 ProPlus for faculty' for $($user)."  -ForegroundColor Magenta
        $O365All = New-MsolLicenseOptions -AccountSkuId regisschool:OFFICESUBSCRIPTION_FACULTY
        Set-MsolUserLicense -UserPrincipalName $user -LicenseOptions $O365All
        Write-Host "Completed, enabled all of  'Office 365 ProPlus for faculty' for $($user)."  -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to enable 'Office 365 ProPlus for faculty' for $($user)." -ForegroundColor Red
   }
}
#Updates Office365 ProPlus for Students, all licenses enabled for a specific user.
function updateAllProPlusStudentUsr($user){
   try{
        Write-Host "Enabling all of  'Office 365 ProPlus for students' for $($user)." -ForegroundColor Magenta
        $O365All = New-MsolLicenseOptions -AccountSkuId regisschool:OFFICESUBSCRIPTION_STUDENT
        Set-MsolUserLicense -UserPrincipalName $user -LicenseOptions $O365All
        Write-Host "Completed, Enabed all of  'Office 365 ProPlus for students' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to enable 'Office 365 ProPlus for students' for $($user)." -ForegroundColor Red
   }
}
#Updates Office365 for faculty education, all licenses enabled for a specific user.
function updateAllEduFacultyUsr($user){
   try{
        Write-Host "Enabling all of  'Office 365 Education for faculty' for $($user)." -ForegroundColor Magenta
        $O365All = New-MsolLicenseOptions -AccountSkuId regisschool:STANDARDWOFFPACK_FACULTY
        Set-MsolUserLicense -UserPrincipalName $user -LicenseOptions $O365All
        Write-Host "Completed, enabled all of  'Office 365 Education for faculty' for $($user)." -ForegroundColor Magenta
      }
   catch{
        Write-Host "Error: failed to enable 'Office 365 Education for faculty' for $($user)." -ForegroundColor Red
   }
 }
#Updates Office365 for student education, all licenses enabled for a specific user.
function updateAllEduStudentUsr($user){
   try{
       Write-Host "Enabling all of  'Office 365 Education for students' for $($user)." -ForegroundColor Magenta
       $O365All = New-MsolLicenseOptions -AccountSkuId regisschool:STANDARDWOFFPACK_STUDENT
       Set-MsolUserLicense -UserPrincipalName $user -LicenseOptions $O365All
       Write-Host "Completed, enabled all of  'Office 365 Education for students' for $($user)." -ForegroundColor Magenta
      }
   catch{
        Write-Host "Error: failed to enable 'Office 365 Education for students' for $($user)." -ForegroundColor Red
   }
}

#Add Office365 ProPlus for faculty, all licenses enabled for a specific user.
function addAllProPlusFacultyUsr($user){
  try{
       Write-Host "Adding all of 'Office 365 ProPlus for faculty' for $($user)." -ForegroundColor Magenta
       Set-MsolUser -UserPrincipalName $user -UsageLocation GB  
       $O365AllFac = New-MsolLicenseOptions -AccountSkuId regisschool:OFFICESUBSCRIPTION_FACULTY
       Set-MsolUserLicense -UserPrincipalName $user -AddLicenses regisschool:OFFICESUBSCRIPTION_FACULTY -LicenseOptions $O365AllFac  
       Write-Host "Completed, added all of  'Office 365 ProPlus for faculty' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to add 'Office 365 ProPlus for faculty' for $($user)." -ForegroundColor Red
   }
}
#Add Office365 ProPlus for Students, all licenses enabled for a specific user.
function addAllProPlusStudentsUsr($user){
    try{
       Write-Host "Adding all of 'Office 365 ProPlus for students' for $($user)." -ForegroundColor Magenta
       Set-MsolUser -UserPrincipalName $user -UsageLocation GB
       $addO365AllStu = New-MsolLicenseOptions -AccountSkuId regisschool:OFFICESUBSCRIPTION_STUDENT
       Set-MsolUserLicense -UserPrincipalName $user -AddLicenses regisschool:OFFICESUBSCRIPTION_STUDENT -LicenseOptions $addO365AllStu
       Write-Host "Completed, added all of 'Office 365 ProPlus for students' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to add 'Office 365 ProPlus for students' for $($user)." -ForegroundColor Red
   }
}
#Add Office365 for faculty education, all licenses enabled for a specific user.
function addAllEduFacultyUsr($user){
  try{
        Write-Host "Adding all of  'Office 365 Education for faculty' for $($user)." -ForegroundColor Magenta
        Set-MsolUser -UserPrincipalName $user -UsageLocation GB
        $O365All = New-MsolLicenseOptions -AccountSkuId regisschool:STANDARDWOFFPACK_FACULTY
        Set-MsolUserLicense -UserPrincipalName $user -AddLicenses regisschool:STANDARDWOFFPACK_FACULTY -LicenseOptions $O365All
        Write-Host "Completed, added all of  'Office 365 Education for faculty' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to add 'Office 365 Education for students' for $($user)." -ForegroundColor Red
   }
}
#Add Office365 for student education, all licenses enabled for a specific user.
function addAllEduStudentUsr($user){
    try{
       Write-Host "Adding all of  'Office 365 Education for students' for $($user)." -ForegroundColor Magenta
       Set-MsolUser -UserPrincipalName $user -UsageLocation GB
       $O365All = New-MsolLicenseOptions -AccountSkuId regisschool:STANDARDWOFFPACK_STUDENT
       Set-MsolUserLicense -UserPrincipalName $user -AddLicenses  regisschool:STANDARDWOFFPACK_STUDENT -LicenseOptions $O365All
       Write-Host "Completed, added all of  'Office 365 Education for students' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to add 'Office 365 Education for students' for $($user)." -ForegroundColor Red
   }
}


#Add sellected licenses under the SKU of Office365 ProPlus for faculty.
function addSellectedProPlusFacultyUsr($user){
    try{
       Write-Host "Adding sellected plans from 'Office 365 ProPlus for faculty' for $($user)." -ForegroundColor Magenta
       Set-MsolUser -UserPrincipalName $user -UsageLocation GB  
       $O365AllFac = New-MsolLicenseOptions -AccountSkuId regisschool:OFFICESUBSCRIPTION_FACULTY -DisabledPlans SWAY
       Set-MsolUserLicense -UserPrincipalName $user -AddLicenses regisschool:OFFICESUBSCRIPTION_FACULTY -LicenseOptions $O365AllFac  
       Write-Host "Completed, added sellected plans from 'Office 365 ProPlus for faculty' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: Adding sellected plans from 'Office 365 ProPlus for faculty' for $($user)." -ForegroundColor Red
   }
}
#Add sellected licenses under the SKU of Office365 ProPlus for students.
function addSellectedProPlusStudentsUsr($user){
    try{
       Write-Host "Adding sellected plans from 'Office 365 ProPlus for students' for $($user)." -ForegroundColor Magenta
       Set-MsolUser -UserPrincipalName $user -UsageLocation GB
       $addO365AllStu = New-MsolLicenseOptions -AccountSkuId regisschool:OFFICESUBSCRIPTION_STUDENT -DisabledPlans SWAY, OFFICE_FORMS_PLAN_2
       Set-MsolUserLicense -UserPrincipalName $user -AddLicenses regisschool:OFFICESUBSCRIPTION_STUDENT -LicenseOptions $addO365AllStu
       Write-Host "Completed, adding selected plans for 'Office 365 ProPlus for students' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to add sellected plans for 'Office 365 ProPlus for students' for $($user)." -ForegroundColor Red
   }
}
#Add sellected licenses under the Office365 for faculty education.
function addSellectedEduFacultyUsr($user){
  try{
        Write-Host "Adding all of 'Office 365 Education for faculty' for $($user)." -ForegroundColor Magenta
        Set-MsolUser -UserPrincipalName $user -UsageLocation GB
        $O365All = New-MsolLicenseOptions -AccountSkuId regisschool:STANDARDWOFFPACK_FACULTY -DisabledPlans SCHOOL_DATA_SYNC_P1, STREAM_O365_E3, TEAMS1, POWERAPPS_O365_P2, RMS_S_ENTERPRISE, YAMMER_EDU

        Set-MsolUserLicense -UserPrincipalName $user -AddLicenses regisschool:STANDARDWOFFPACK_FACULTY -LicenseOptions $O365All
        Write-Host "Completed, added all of  'Office 365 Education for faculty' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to add 'Office 365 Education for students' for $($user)." -ForegroundColor Red
   }
}
#Add sellected licenses under the Office365 for student education.
function addSellectedEdcuStudentUsr($user){
    try{
       Write-Host "Adding all of  'Office 365 Education for students' for $($user)." -ForegroundColor Magenta
       Set-MsolUser -UserPrincipalName $user -UsageLocation GB
       $O365All = New-MsolLicenseOptions -AccountSkuId regisschool:STANDARDWOFFPACK_STUDENT -DisabledPlans AAD_BASIC_EDU, SCHOOL_DATA_SYNC_P1, STREAM_O365_E3, TEAMS1, Deskless, FLOW_O365_P2, POWERAPPS_O365_P2, RMS_S_ENTERPRISE, YAMMER_EDU, MCOSTANDARD           
       Set-MsolUserLicense -UserPrincipalName $user -AddLicenses  regisschool:STANDARDWOFFPACK_STUDENT -LicenseOptions $O365All 
       Write-Host "Completed, added all of  'Office 365 Education for students' for $($user)." -ForegroundColor Magenta
   }
   catch{
        Write-Host "Error: failed to add 'Office 365 Education for students' for $($user)." -ForegroundColor Red
   }
}


#Imports the email addresses from the csv file and edit their licenses.
function importCSVUsrsAndEditLicenses($importPath){
    try{
        Write-Host "Importing CSV file: $($importPath)" -ForegroundColor Magenta
        #Imports from the CSV file path given.
        Import-Csv $importPath | ForEach-Object {
        
            #Stores the UserPrincipalName
            $email = $_.UserPrincipalName        
         
            Write-Host "Editing $($email)" -ForegroundColor Magenta
         
            #Check which function user wants to use.
            #Asigns all licenses under the SKU of Office365 ProPlus for faculty.
            if($updateAllProPlusFacultyUser.Contains("csv")){               
                updateAllProPlusFacultyUsr($email)
            }
            #Asigns all licenses under the SKU of Office365 ProPlus for students.
            elseif($updateAllProPlusStudentsUser.Contains("csv")){
               updateAllProPlusStudentUsr($email)
            }
            #Assign all licenses under the SKU of Office 365 Education for faculty.
            elseif($updateAllEduFacultyUser.Contains("csv")){
                updateAllEduFacultyUsr($email)
            }
            #Assign all licenses under the SKU of Office 365 Education for students.
            elseif($updateAllEduStudentUser.Contains("csv")){
                 updateAllEduStudentUsr($email)
            }

            #Add all licenses under the SKU of Office365 ProPlus for faculty.
            elseif($addAllProPlusFacultyUser.Contains("csv")){
               addAllProPlusFacultyUsr($email)
            }
            #Add all licenses under the SKU of Office365 ProPlus for students.
            elseif($addAllProPlusStudentsUser.Contains("csv")){
               addAllProPlusStudentsUsr($email)
            }
            #Add all licenses under the SKU of Office 365 Education for faculty.
            elseif($addAllEduFacultyUser.Contains("csv")){
               addAllEduFacultyUsr($email)
            }
            #Add all licenses under the SKU of Office 365 Education for students.
            elseif($addAllEduStudentUser.Contains("csv")){
               addAllEduStudentUsr($email)
            }

             #Add sellected licenses under the SKU of Office365 ProPlus for faculty.
            elseif($addSellectedProPlusFacultyUser.Contains("csv")){
                addSellectedProPlusFacultyUsr($email)
            }
            #Add sellected licenses under the SKU of Office365 ProPlus for students.
            elseif($addSellectedProPlusStudentsUser.Contains("csv")){
               addSellectedProPlusStudentsUsr($email)
            }
            #Add sellected licenses under the SKU of Office 365 Education for faculty.
            elseif($addSellectedEduFacultyUser.Contains("csv")){
               addSellectedEduFacultyUsr($email)
            }  
            #Add sellected licenses under the SKU of Office 365 Education for students.
            elseif($addSellectedEdcuStudentUser.Contains("csv")){
                addSellectedEdcuStudentUsr($email)
            }
            #If none of the above is selected, produces an error.
            else{
                Write-Host "No Extra parameter sellected, please check your code." -ForegroundColor Red
            }
        }
         Write-Host "Finished importing CSV file: $($importPath)" -ForegroundColor Magenta
    }
    catch{
        Write-Host "Failed: Could not import CSV file from $($importPath)" -ForegroundColor Red
    }

}

#Clear the screen.
cls