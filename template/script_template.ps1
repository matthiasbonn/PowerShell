#requires -version 2
<#
.SYNOPSIS
  <Overview of script>
 
.DESCRIPTION
  <Brief description of script>
 
.PARAMETER <Parameter_Name>
    <Brief description of parameter input required. Repeat this attribute if required>
 
.INPUTS
  <Inputs if any, otherwise state None>
 
.OUTPUTS
  <Outputs if any, otherwise state None - example: Log file stored in C:\Windows\Temp\<name>.log>
 
.NOTES
  Version:        1.0
  Author:         Bonn, Matthias - Alegri International Service GmbH
  Creation Date:  <Date>
  Purpose/Change: Initial script development
  
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>
[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]
param (
		[Parameter(Mandatory=$False)]
		[string[]]$Log,
        [Parameter(Mandatory=$False)]
	    [Alias('debugger')]
		[string[]]$DGBPath
	)
#region Initialization code
    Write-Verbose 'Initialize stuff in Begin block'
    #region Modules / Addins / DotSourcing
     #if (-not(Get-PSSnapin Quest.ActiveRoles.ADManagement -ErrorAction Silentlycontinue)){Add-PSSnapin Quest.ActiveRoles.ADManagement}
    #endregion Modules / Addins / DotSourcing
    
    #Dump all existing Variables in a variable for housekeeping
    $startupVariables=''
    new-variable -force -name startupVariables -value ( Get-Variable | % { $_.Name } )
    $TemplateVersion = 0.2
    $ScriptVersion   = 0.1
    $ScriptPath      = (Split-Path ((Get-Variable MyInvocation).Value).MyCommand.Path)
    $ScriptName      = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName)
    If (!$Log){
        $LogPath     = "$ScriptPath\Logs"
        }
    Else
    {
        $LogPath     = $Log
    }
    $LogFile         = "$Logpath\$Scriptname.log"
    $DateFormat      = Get-Date -Format 'yyyyMMdd_HHmmss'
    Write-Verbose "Start Script $ScriptName V$ScriptVersion at $DateFormat from $ScriptPath" -verbose
    IF(!(Test-Path $LogPath)) {mkdir $LogPath}

    #region nlog

  <# Possible Usage for Logging
        $PSlogger.Debug("Debug Message") 
        $PSlogger.Info("Info Message") 
        $PSlogger.Warn("Warn Message") 
        $PSlogger.Error("Error Message") 
        $PSlogger.Trace("Trace Message")  
        $PSlogger.Fatal("Fatal Message") 
  #> 

    # configuration for the nlog feature
    [Reflection.Assembly]::LoadFile(“$scriptPath\NLog.dll”)

    # Load from Nlog.config file
    $nlogconfig = new-object NLog.Config.XmlLoggingConfiguration("$ScriptPath\NLog.config")

    # Assign configuration
    ([NLog.LogManager]::Configuration)=$nlogconfig

    # Change filename
    [NLog.Targets.FileTarget]$fileTarget = [NLog.Targets.FileTarget]([NLog.LogManager]::Configuration.FindTargetByName('logfile'))
    $fileTarget.FileName = $LogFile
    $PSlogger = [NLog.LogManager]::GetLogger('PSLogger')
    
    #endregion nlog
 
#endregion Initialization code


#region functions

function Check-Powershell64{
 $is64Bit=[Environment]::Is64BitProcess
 return $is64Bit
}

function Cleanup-Variables {

  Get-Variable |

    Where-Object { $startupVariables -notcontains $_.Name } |

   % { try {Remove-Variable -Name "$($_.Name)" -Force -Scope "global"  -ErrorAction SilentlyContinue -WarningAction SilentlyContinue} 
catch{} }

}

 


#endregion fuctions


#region Process data
$PSlogger.Info('Begin of execution')
IF(Check-Powershell64){$PSlogger.Info('Running in 64 BIT Process..')}



#endregion Process data

#region Finalizing 
    
    $PSlogger.Info('Cleanup started ...')
    Cleanup-Variables
#endregion Finalizing 

