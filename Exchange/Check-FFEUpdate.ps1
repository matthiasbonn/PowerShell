#requires -version 2
<#
.SYNOPSIS
  Check it the Fore Front Enginge Pattern is updated on the local Exchange System
 
.DESCRIPTION
  Check the pattern update and log to text file and send a e-mail notification 
 
 .PARAMETER SendMail
	Send Mail after completion. Set to $True to enable. If enabled, -MailFrom, -MailTo, -MailServer are mandatory
	
.PARAMETER MailFrom
	Email address to send from. Passed directly to Send-MailMessage as -From
	
.PARAMETER MailTo
	Email address to send to. Passed directly to Send-MailMessage as -To
	
.PARAMETER MailServer
	SMTP Mail server to attempt to send through. Passed directly to Send-MailMessage as -SmtpServer  

 
.INPUTS
  None
 
.OUTPUTS
  Update the logs\Check-FFEUpdate.log with the status of the check and send an e-mail if configured
 
.NOTES
  Version:        1.0
  Author:         Bonn, Matthias - Alegri International Service GmbH
  Creation Date:  2015-11-24
  Purpose/Change: Initial script development
  
.EXAMPLE
  .\Check-FFEUpdate -SendMail $True -$MailFrom max.mustermann@contoso.com -MailTo matthias.bonn@alegri.eu -MailServer localhost
#>
[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High',DefaultParametersetName="default")]
param (
		[Parameter(ParameterSetName='dafault',Mandatory=$False)]
		[string[]]$Log,
        [parameter(ParameterSetName='SendMail',Mandatory=$true,HelpMessage='SendMail $True/$False to inform via E-Mail')]
        [string]$SendMail,
        [parameter(ParameterSetName='SendMail',Mandatory=$true,HelpMessage='Mail from')]
        [string]$MailFrom,
        [parameter(ParameterSetName='SendMail',Mandatory=$true,HelpMessage='Mail to')]
        [string]$MailTo,
        [parameter(ParameterSetName='SendMail',Mandatory=$true,HelpMessage='Mailserver to be used')]
        [string]$MailServer,
        [Parameter(ParameterSetName='dafault',Mandatory=$False)]
	    [Alias('debugger')]
		[string[]]$DGBPath
	)
#region Initialization code
    Write-Verbose 'Initialize stuff in Begin block'
    #region Modules / Addins / DotSourcing
     #if (-not(Get-PSSnapin Quest.ActiveRoles.ADManagement -ErrorAction Silentlycontinue)){Add-PSSnapin Quest.ActiveRoles.ADManagement}
    if (-not(Get-PSSnapin Microsoft.Forefront.Filtering.management.Powershell -ErrorAction Silentlycontinue)){Add-PSSnapin Microsoft.Forefront.Filtering.management.Powershell}
    #endregion Modules / Addins / DotSourcing
    
    #Dump all existing Variables in a variable for housekeeping
    $startupVariables=''
    new-variable -force -name startupVariables -value ( Get-Variable | % { $_.Name } )
    $TemplateVersion = 0.1
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

    % { Remove-Variable -Name "$($_.Name)" -Force -Scope "global" }

}

 


#endregion fuctions


#region Process data
$PSlogger.Info('Begin of execution')
IF(Check-Powershell64){$PSlogger.Info('Running in 64 BIT Process..')}
$FFEnegine = Get-EngineUpdateInformation
If ($FFEnegine.UpdateStatus -ne 'UpdateAttemptNoUpdate') {
    IF ($FFEnegine.SignatureDateTime -gt ((Get-Date).AddDays(-1))) {
        $PSlogger.Warning("ForeFront Engine outdated | $env:computername")
        If ($SendMail) {
            $PSlogger.Info('Sending E-Mail ...')
            $Output = "FF Engine Update Error on $env:hostname"
            Send-MailMessage -To $MailTo -From $MailFrom -SmtpServer $MailServer -Subject "FF Engine Update Error on $env:computername" -BodyAsHtml $Output
        }
    }
}
Else{
    $PSlogger.Info("ForeFront Enginge up to date | $env:computername")
    If ($SendMail) {
            $PSlogger.Info('Sending E-Mail ...')
            $Output = "FF Engine Update on $env:computername ok"
            Send-MailMessage -To $MailTo -From $MailFrom -SmtpServer $MailServer -Subject "FF Engine Update ok on $env:computername" -BodyAsHtml $Output
    }
}

#endregion Process data

#region Finalizing 
    
    $PSlogger.Info('Cleanup started ...')
    Cleanup-Variables
#endregion Finalizing 
 
