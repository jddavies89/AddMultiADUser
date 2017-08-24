############################################################
##
##Author by Joe Richards
##
##This is a simple logging module example with capabilities to 
##          *Log at startup. 
##          *Log on close.
##          *Delete the log.
##          *Check the log file.
##          *Open the log file.
##          *Add text to the log on user input.
##          *Logs if the module has been loaded. 
##
##############################################################

#Stores the log file details.
param(
    #Storesd the executional path.
    $path = (split-path $script:MyInvocation.MyCommand.Path -parent),
    #Stores the log name.
    $file = "\log.log",
    #Stores the path and the name together.
    $fileDir = ($path + $file)
)#End paramter.

#Start function
function startApp{
    #Creates the log file if it does not exist and adds to the log file.
    Add-Content -Value ("Started application at " + (Get-Date -Format "%d-%M-%y %h:%m:%s")) -Path $fileDir
}#End function

#Start function
function closeApp{
    #Creates the log file if it does not exist and adds to the log file.
    Add-Content -Value ("Closed application at " + (Get-Date -Format "%d-%M-%y %h:%m:%s")) -Path $fileDir
}#End function

#Start function
function delFile{
    #Checks if the log.log file exists.
    if(Test-Path $fileDir){
       #Removes log.log file if it exists.
       Remove-Item -Path $fileDir
    }
}#End function

#Start function
function checkFile{
    #Checks if the log.log file is not in the curent working directory.
    if(!(Test-Path $fileDir)){
        #Creates the log file if it does not exist and adds to the log file.
        Add-Content -Value ("Created log file at " + (Get-Date -Format "%d-%M-%y %h:%m:%s")) -Path $fileDir
    }
    else{
        #Creates the log file if it does not exist and adds to the log file.
        Add-Content -Value ("Log file exists at " + (Get-Date -Format "%d-%M-%y %h:%m:%s")) -Path $fileDir
    }
}#End function

#Start function
function addlog{ 
    #Stores the text that will be appended to the log file.  
    param($text)
    #Appends the text to the file log.log.
    Add-Content -Value $text -Path $fileDir
}#End function

#Start function
function openlog{
    #checks that the file exists.
    if(Test-Path $fileDir){
        #Opens the file log.log
        Invoke-Item $fileDir
    }
    else{
         #Creates the log file if it does not exist and adds to the log file.
         Add-Content -Value ("Couldn't open the log file at "  + (Get-Date -Format "%d-%M-%y %h:%m:%s")) -Path $fileDir
    }
}#End function

#Start function
function loadModuleLog{
    #Creates the log file if it does not exist and adds to the log file.
    Add-Content -Value ("Module has been loaded at " + (Get-Date -Format "%d-%M-%y %h:%m:%s")) -Path $fileDir
}