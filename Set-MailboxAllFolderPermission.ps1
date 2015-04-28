#requires -version 2
<#
    .SYNOPSIS
    Set folder permissions on all folders that conaitns the foldername in the mailbox UserAlias for a given user 
  
    .DESCRIPTION
    <Brief description of script>
 
    .PARAMETER UserAlias
    Mailbox in that the folderpermission will be set, the identity an be the alias or the smtpt address

    .PARAMETER GrantToAlias
    The user that will be have AccesRights in the mailbox of the UserAlias.   
 
    .PARAMETER Foldername
    The Folder what would be change.It can be a part of the foldername or / 
 
    .PARAMETER AccessRight
    Which AccesRights should bes set on the folders. It must be in 'None','Owner','PublishingEditor','Editor','PublishingAuthor','Author','NonEditingAuthor','Reviewer','Contributor'
    Default is Author 

    .INPUTS
    NONE
 
    .OUTPUTS
    NONE
 
    .NOTES
    Version:        1.0
    Author:         Bonn, Matthias - matthias.bonn@mb-itconsult.de
    Creation Date:  28.04.2015
    Purpose/Change: Initial script development
  
    .EXAMPLE
    .\Set_MailboxAllFolderPermission -UserAlias MaxMuster -GrantToAlias SabineMuster
    .\Set_MailboxAllFolderPermission -UserAlias MaxMuster -GrantToAlias SabineMuster -FolderName "/" -AccessRight Editor
#>
[CmdletBinding(SupportsShouldProcess=$true)]

Param(
    [Parameter(ParameterSetName='AddRights',Mandatory=$True,Position=1)]
    [Parameter(ParameterSetName='RemoveRights',Mandatory=$True,Position=1)]
    [string]$UserAlias,
    [Parameter(ParameterSetName='AddRights',Mandatory=$True,Position=2)]
    [Parameter(ParameterSetName='RemoveRights',Mandatory=$True,Position=2)]
    [string]$GrantToAlias,
    [Parameter(ParameterSetName='AddRights',Mandatory=$False)]
    [string]$FolderName,
    [Parameter(ParameterSetName='AddRights',Mandatory=$False)]
    [ValidateSet('None','Owner','PublishingEditor','Editor','PublishingAuthor','Author','NonEditingAuthor','Reviewer','Contributor')] 
    [alias('Perm')]
    [string]$AccessRight = 'Author',
    [Parameter(ParameterSetName='RemoveRights',Mandatory=$False)]
    [switch]$RemoveRight
)

 
#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
 


#--- CONFIG ---#
#region Configuration
# Script Path/Directories
$ScriptPath   = (Split-Path ((Get-Variable MyInvocation).Value).MyCommand.Path)
 
#endregion configuration


#region declaration
$ScriptVersion = '1.0'

$exclusions = @("/Sync Issues",
    "/Sync Issues/Conflicts",
    "/Sync Issues/Local Failures",
    "/Sync Issues/Server Failures",
    "/Recoverable Items",
    "/Deletions",
    "/Purges",
    "/Versions"
)
#endregion declaration

#region functions 
#-----------------------------------------------------------[Functions]------------------------------------------------------------

<#
    Requires dbgview.exe from Sysinternals to be in your path or 
    modify the function. You should set up a filter in dbview.exe
    before using this to filter on the Category.
#>
 
Function Debug-Message 
{
    [cmdletbinding()]
    Param(
        [Parameter(Position=0,Mandatory=$True,HelpMessage='Enter a message')]
        [string]$Message,
        [string]$Category='PS Trace'
    )
 
    #only run if $TraceEnabled is True
    if ($script:TraceEnabled) 
    {
        #test if dbgview.exe is already running and start it if not.
        if (-NOT (Get-Process -Name Dbgview -ErrorAction SilentlyContinue)) 
        {
            Try 
            {
                #start with /f to skip filter confirmation prompt
                Start-Process G:\Sysinternals\Dbgview.exe /f
                #give the application to start
                Start-Sleep -Seconds 1
            }
            Catch 
            {
                Write-Warning 'Failed to find or start dbgview.exe'
                Return
            }
        } #if dbgview is not running
 
        #display the message in dbgview.exe
        [System.Diagnostics.Debug]::WriteLine($Message,$Category)
    } #if $TraceEnabled
} #close Debug-Message
 
#endregion functions


if($RemoveRight)
{
    # Get-MailboxFolderStatictics replace the /-Char  with cahr 63743
    $fname = $UserAlias+':' + $f.FolderPath.Replace('/','\').Replace([char]63743,'/')
    if ($PSCmdlet.ShouldProcess($fname))
    {
        Remove-MailboxFolderPermission -Identity $fname -User $GrantToAlias -Confirm:$False
    }
    Else
    {
        Remove-MailboxFolderPermission -Identity $fname -User $GrantToAlias -Confirm:$False -WhatIF
    } 

    Remove-MailboxFolderPermission -Identity $fname -User $GrantToAlias
}
Else 
{
    ForEach($f in (Get-MailboxFolderStatistics -Identity $UserAlias | Where-Object {
                $_.FolderPath.Contains($FolderName) -eq $True
            }
    ) ) 
    {
        # Get-MailboxFolderStatictics replace the /-Char  with cahr 63743
        $fname = $UserAlias+':' + $f.FolderPath.Replace('/','\').Replace([char]63743,'/')
        if ($PSCmdlet.ShouldProcess($fname))
        {
            Add-MailboxFolderPermission -identity $fname -User $GrantToAlias -AccessRights $AccessRight
        }
        Else
        {
            Add-MailboxFolderPermission -identity $fname -User $GrantToAlias -AccessRights $AccessRight -WhatIF
        } 
    }
}
