<#
.Synopsis
   This is an advanced function logging module that logs to a log file of the location specified by the user.
.DESCRIPTION
    See logger.ps1 on how to implement this advanced function module.
.PARAMETER Path
    The directory of where the log file is located.
.PARAMTER fileName
    Stores the log file name.
.PARAMTER SimpleLogging
    Logs only the user input and not the script activity.
.PARAMETER AdvancedLogging
    Logs the app activity as well as the logging of the user input.
.PARAMTER Text
    Logs the user input to the text file.
.EXAMPLE
    Log-ToFile -Path C:\logging -fileName logger.log -SimpleLogging -Text hi
    Logs the filename logger.log in C:\logging with only the text of the user input, if C:\logging doesnt exit, it throw an error.
.EXAMPLE
    Log-ToFile -Path C:\logging -fileName logger.log -AdvancedLogging -Text hi
    Logs the filename logger.log in C:\logging with the activity of the module and text from the user input, if C:\logging doesnt exit, it throw an error.
.EXAMPLE
    Log-ToFile -Path C:\logging -fileName logger.log -AdvancedLogging -Text "hi im text."
    Logs the filename logger.log in C:\logging with the acitvity of the module and text from the user input, if C:\logging doesnt exit, it throw an error.
.Notes
   Author: Joe Richards
   Date:   09/02/2017
.LINK
  https://github.com/joer89/AddMultiADUser.git
#>


function Log-ToFile {

[CmdletBinding(SupportsShouldProcess=$true)]        

    Param (        
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [ValidateScript({
            if(Test-Path $_){$true}else{Throw "Invalid path given: $_"}
            })]
        [string]$Path,

        [string]$fileName,

        [switch]$SimpleLogging,

        [switch]$AdvancedLogging,

        [string]$Text
    )

    Begin{        
        if($AdvancedLogging){ 
            #Creates the log file if it does not exist and adds to the log file.
            Add-Content -Value ("Module has been loaded at " + (Get-Date -Format "%d-%M-%y %h:%m:%s")) -Path (Join-Path $Path -ChildPath "\$($fileName)")
            #Creates the log file if it does not exist and adds to the log file.
            Add-Content -Value ("Started application at " + (Get-Date -Format "%d-%M-%y %h:%m:%s")) -Path (Join-Path $Path -ChildPath "\$($fileName)")   
        }
    }
    Process{
        if(($SimpleLogging) -or ($AdvancedLogging)){        
            #Creates the log file if it does not exist and adds to the log file.
            Add-Content -Value ($Text + (Get-Date -Format " %d-%M-%y %h:%m:%s")) -Path (Join-Path $Path -ChildPath "\$($fileName)")   
        }
    }
    End{
        if($AdvancedLogging){
            #Creates the log file if it does not exist and adds to the log file.
            Add-Content -Value ("Closed application at " + (Get-Date -Format "%d-%M-%y %h:%m:%s")) -Path (Join-Path $Path -ChildPath "\$($fileName)")   
        }
    }
}
