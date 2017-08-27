<#
.Synopsis
   This script checks each username in Users.csv to see if any already exist in Active Directory. 
.Notes
   Author: Joe Richards
   Date:   24/08/2017
.LINK
  https://github.com/joer89/AddMultiADUser.git
#>

param(
    #Stores the array list of duplicate names and free names.
    $duplicateNames = (New-Object System.Collections.ArrayList),
    $names = (New-Object System.Collections.ArrayList)
)

#Imports the Active Directory module for conencting to AD.
Import-Module ActiveDirectory -ErrorAction Stop

#Check the names of .\Users.csv to search Active Directory.
function check_ADName(){
    #Imports the Users csv file and compares each row.
    Import-Csv .\Users.csv | ForEach-Object {       
        #Stores the username.
        $Username = $_.Username
        #Stores the searched username.
        $User = Get-ADUser -LDAPFilter "(sAMAccountName=$Username)"
        #Checks to see if there is a match.
        if ($Username -eq (Get-ADUser -Filter 'Name -like "$($User)"' | FT Name -A)){
            #If there is a match, Write to screen and store the username in the array $duplicateNames.
            Write-Host "Duplicate username in Active Directory >>> $($Username)" -ForegroundColor Red
            #Stores the username in the array duplicateNames.
            $duplicateNames.Add($Username)
        }
        else{
            #If there is no match, prints out on the screen no match.
            Write-Host "Scanned $($Username) >>> not in Active Directory." -ForegroundColor Magenta
            #Stores the no matched username in the array.
            $names.Add($Username)
        }
    }
}

#Display the results.
function Results{
    #Displays each duplicate name on screen.
    Write-Host "Duplicate names:" -ForegroundColor Red
    foreach ($DupName in $duplicateNames){
        Write-Host $DupName -ForegroundColor Red
    }

    #Displays the count in the array of duplicate names.
    Write-Host "Duplicant name count is $($duplicateNames.Count)" -ForegroundColor Red
    #Displays the count in the array of names.
    Write-Host "Free names count is $($names.Count)" -ForegroundColor Magenta
}

#Run the two functions.
check_ADName
Results
