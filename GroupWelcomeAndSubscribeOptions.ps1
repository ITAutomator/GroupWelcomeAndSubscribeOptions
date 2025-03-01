<#
.SYNOPSIS
    Sets AutoSubscribeNewMembers and WelcomeMessageEnabled for a specified Microsoft 365 Group,
    then applies subscription changes to existing members to match the desired setting.

.DESCRIPTION
    1. Checks if Connect-ExchangeOnline command is available (i.e., ExchangeOnlineManagement module is installed).
       If not, prints instructions, prompts to press a key, and exits.
    2. Checks if you're connected to Exchange Online by calling Get-OrganizationConfig.
       If it fails, prompts to press a key, then exits.
       If successful, displays which tenant/domain you are connected to.
    3. Prompts the user for a group name. If no name is entered, prompts to press a key, then exits.
    4. Prompts the user whether new members should be auto-subscribed (True/False).
       If invalid input, prompts to press a key, then exits.
    5. Prompts the user whether welcome messages should be suppressed (True/False).
       If invalid input, prompts to press a key, then exits.
    6. Retrieves the group and compares current group settings to the desired settings.
       If it fails to retrieve, prompts to press a key, then exits.
    7. If needed, updates the group's AutoSubscribeNewMembers and WelcomeMessageEnabled properties.
       If it fails, prompts to press a key, then exits.
    8. Retrieves all existing members and all current subscribers for the group.
       If it fails, prompts to press a key, then exits.
    9. For each member, checks if they match the desired subscription state:
       - If AutoSubscribeNewMembers = True, ensure the user is subscribed.
       - If AutoSubscribeNewMembers = False, ensure the user is unsubscribed.
    10. Outputs whether each user was already correct or if it’s being updated, and completes.

.NOTES
    - If your environment still uses -WelcomeEmailDisabled instead, replace or remove
      the -WelcomeMessageEnabled part accordingly.
#>
######################
## Main Procedure
######################

#################### Transcript Open
$Transcript = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $Transcript | Out-Null
#################### Transcript Open

###
## To enable scrips, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
###
### Main function header - Put ITAutomator.psm1 in same folder as script
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
$scriptVer      = "v"+(Get-Item $scriptFullname).LastWriteTime.ToString("yyyy-MM-dd")
$psm1="$($scriptDir)\ITAutomator.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
# Get-Command -module ITAutomator  ##Shows a list of available functions
######################

#######################
## Main Procedure Start
#######################
Write-Host "-----------------------------------------------------------------------------"
Write-Host "$($scriptName) $($scriptVer)       Computer:$($env:computername) User:$($env:username) PSver:$($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
Write-Host ""
Write-Host " Set M365 Group Auto-Subscription & Welcome Email."
Write-Host ""
Write-Host " M365 Groups are  here: https://admin.microsoft.com/#/groups/_/UnifiedGroup"
Write-host "             (dynamic): https://admin.microsoft.com/#/groups/_/UnifiedGroup/DynamicMembership"
Write-Host "-----------------------------------------------------------------------------"
PressEnterToContinue

# Check if Connect-ExchangeOnline is available
if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
    Write-Host "ERROR: 'Connect-ExchangeOnline' command was not found."
    Write-Host "Please install the ExchangeOnlineManagement module using:"
    Write-Host "   Install-Module ExchangeOnlineManagement"
    Write-Host "Or load the module if it is already installed, then try again."
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    return
}

# Check if we are already connected to Exchange Online
while ($true) {
    try {
        $orgConfig = Get-OrganizationConfig -ErrorAction Stop
        # The Identity property typically shows your tenant's name or domain
        $tenantNameOrDomain = $orgConfig.Identity
        Write-Host "You are currently connected to tenant: " -NoNewline
        Write-host $tenantNameOrDomain -ForegroundColor Green
        $response = AskForChoice "Choice: " -Choices @("&Use this connection","&Disconnect and try again","E&xit") -ReturnString
        # If the user types 'exit', break out of the loop
        if ($response -eq 'Disconnect and try again') {
            Write-Host "Disconnect-ExchangeOnline..."
            $null = Disconnect-ExchangeOnline -Confirm:$false
            PressEnterToContinue "Done. Press <Enter> to connect again."
            Continue # loop again
        }
        elseif ($response -eq 'exit') {
            return
        }
        else { # on to next step
            break
        }
    }
    catch {
        Write-Host "ERROR: Not connected to Exchange Online or invalid session."
        Write-Host "We will try 'Connect-ExchangeOnline' to authenticate. Before we do, open a browser to an admin session on the desired tenant."
        Write-Host "Press any key to continue..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Write-Host "Connect-ExchangeOnline ... " -ForegroundColor Yellow
        Connect-ExchangeOnline -UseWebLogin -ShowBanner:$false
        Write-Host "Done" -ForegroundColor Yellow
        Continue # loop again
    }
}
Write-Host

while ($true)
{
    # Prompt for group name
    $groupNames_str = Read-Host "Enter the M365 Group name or email address, separate multiple groups with commas to loop (blank to exit)"
    if ([string]::IsNullOrWhiteSpace($groupNames_str)) {
        Write-Host "No group name entered. Exiting script."
        Start-Sleep 2
        Break
    }
    # Loop through multiple groups separated by commas
    $changes_made_total = 0
    $groupNames = @($groupNames_str.Split(","))
    ForEach ($groupName in $groupNames)
    {
        $groupName = $groupName.Trim()
        $changes_made = 0
        # Retrieve the group
        try {
            $group = Get-UnifiedGroup -Identity $groupName -ErrorAction Stop
        } catch {
            Write-Host "ERROR: Failed to retrieve the group '$groupName'."
            PressEnterToContinue "Press Enter to contine."
            Continue 
        }
        # Begin changes
        Write-Host "                                      Group settings: " -NoNewline
        Write-host $group.DisplayName -ForegroundColor Yellow -NoNewline
        Write-Host " (The recommended settings are True for both)"
        Write-Host "  Suppress welcome messages (!WelcomeMessageEnabled): " -NoNewline
        write-host (-not $group.WelcomeMessageEnabled) -ForegroundColor Yellow
        Write-Host "                            Subscribe future members: " -NoNewline
        write-host $group.AutoSubscribeNewMembers -ForegroundColor Yellow
        foreach ($grp_setting in @("Welcome","Subscribe"))
        {
            if ($grp_setting -eq "Welcome") {
                $current = $group.WelcomeMessageEnabled
                $setprop =      "UnifiedGroupWelcomeMessageEnabled"
                $target = AskForChoice "Suppress welcome messages?" -Choices @("&True","&False","&Dont adjust") -ReturnString
            }
            elseif ($grp_setting  -eq "Subscribe") {
                $current = $group.AutoSubscribeNewMembers
                $setprop =       "AutoSubscribeNewMembers"
                $target = AskForChoice "Subscribe future members (so they get msgs in their inbox)?" -Choices @("&True","&False","&Dont adjust") -ReturnString
            }
            if ($target -eq "Dont adjust") {
                Write-Host "Group setting [$($setprop)] not adjusted from [$($current)]"
            } # don't adjust
            else {
                if ($grp_setting -eq "Welcome") {
                    $target_bool = -not [bool]$target # requires a flip
                } # welcome
                elseif ($grp_setting  -eq "Subscribe") {
                    $target_bool = [bool]$target
                } # subscribe
                # adjust group setting if needed
                if ($current -ne $target_bool) {
                    $Parameters = @{
                        Identity = $groupName
                        $setprop = $target_bool 
                    }
                    Set-UnifiedGroup @Parameters
                    Write-Host "Group setting [$($setprop)] adjusted from [$($current)] to " -NoNewline
                    Write-Host "[$($target_bool)]" -ForegroundColor Yellow
                    $changes_made +=1
                } else {
                    Write-Host "Group setting already OK." -ForegroundColor Green
                } # no group setting adjustment needed
            } # adjust
        } # two grp_settings
        # user changes?
        Write-host "---------------------------------------------------------------------------------------------------------------"
        Write-host "Subscribe existing members" -ForegroundColor Yellow
        Write-Host "This next setting adjusts how users have already elected to be subsribed (true) or not subscribed (false)"
        Write-Host "True means they will get a copy of group messages in their personal Inbox"
        Write-Host "False means they won't (the group collects the mail regardless - This is the default)"
        Write-Host "Caution: This will overwrite user selections (Use Dont adjust to check current selections)"  -ForegroundColor Yellow
        Write-host "---------------------------------------------------------------------------------------------------------------"
        $target_users = AskForChoice "Subscribe existing members? (Recommended setting is True)" -Choices @("&True","&False","&Dont adjust (just check)") -ReturnString
        Write-Host "Retrieving [$($groupName)] existing members and subscribers ..."
        try {
            $groupMembers     = Get-UnifiedGroupLinks -Identity $groupName -LinkType Members     -ErrorAction Stop
            $groupSubscribers = Get-UnifiedGroupLinks -Identity $groupName -LinkType Subscribers -ErrorAction Stop
        } catch {
            Write-Host "ERROR: Failed to retrieve group links. $_"
            PressEnterToContinue
            Continue
        }
        Write-Host "  [$($groupName)] members : " -NoNewline
        Write-host $groupMembers.count -ForegroundColor Green
        Write-Host "  [$($groupName)] subscribers : " -NoNewline
        Write-host $groupSubscribers.count -ForegroundColor Green
        # Reconcile each member's subscription status
        $count_users = 0
        if ( ($target_users -eq "True") -or ($target_users -eq "Dont adjust (just check)")) {
            # If AutoSubscribeNewMembers = True, ensure everyone is subscribed
            Write-Host "Checking each member subscribed status ..."
            foreach ($groupMember in $groupMembers) {
                $count_users +=1
                Write-Host "  [$($count_users) of $($groupMembers.count)] $($groupMember.DisplayName) " -NoNewline
                if ($groupSubscribers.Guid -contains $groupMember.Guid) {
                    Write-Host "OK: IsSubscribed already"  -ForegroundColor Green
                } else {
                    if ($target_users -eq "Dont adjust (just check)") {
                        Write-Host "OK: NotSubscribed"  -ForegroundColor Green
                    } # not adjust
                    else {
                        Write-Host "Changed: Subscribing"  -ForegroundColor Yellow
                        try {
                            Add-UnifiedGroupLinks -Identity $groupName -LinkType Subscribers -Links $groupMember.Guid -ErrorAction Stop
                            $changes_made +=1
                        } catch {
                            Write-Host "    ERROR: Could not subscribe $($groupMember.DisplayName)  . $_"
                            PressEnterToContinue
                        }
                    } # adjust
                }
            }
        } else {
            # If AutoSubscribeNewMembers = False, ensure no members are subscribed
            Write-Host "Ensuring each member is unsubscribed..."
            foreach ($groupMember in $groupMembers) {
                $count_users +=1
                Write-Host "  [$($count_users) of $($groupMembers.count)] $($groupMember.DisplayName) " -NoNewline
                if ($groupSubscribers.Guid -contains $groupMember.Guid) {
                    Write-Host "Changed: Unsubscribing"  -ForegroundColor Yellow
                    try {
                        Remove-UnifiedGroupLinks -Identity $groupName -LinkType Subscribers -Links $groupMember.Guid -ErrorAction Stop
                        $changes_made +=1
                    } catch {
                        Write-Host "    ERROR: Could not unsubscribe $($groupMember.DisplayName). $_"
                        PressEnterToContinue
                    }
                } else {
                    Write-Host "OK: already not subscribed"  -ForegroundColor Green
                }
            }
        }
        Write-Host "Done with [$($groupName)]. Changes made in this group: " -NoNewline
        Write-host $changes_made -ForegroundColor Green
        Write-Host "-----------------------------------------------------------------------------"
        $changes_made_total += $changes_made
        PressEnterToContinue "Press <Enter> for next group"
    } # each $groupName
    Write-host "Done. Changes made in all groups: " -NoNewline
    Write-host $changes_made_total -ForegroundColor Green
    Write-Host "-----------------------------------------------------------------------------"
    PressEnterToContinue "The program will start again in case you want to change other groups (Press enter at the next prompt to exit)."
    Write-Host "-----------------------------------------------------------------------------"
}
#################### Transcript Save
Stop-Transcript | Out-Null
$date = get-date -format "yyyy-MM-dd_HH-mm-ss"
New-Item -Path (Join-Path (Split-Path $scriptFullname -Parent) ("\Logs")) -ItemType Directory -Force | Out-Null #Make Logs folder
$TranscriptTarget = Join-Path (Split-Path $scriptFullname -Parent) ("Logs\"+[System.IO.Path]::GetFileNameWithoutExtension($scriptFullname)+"_"+$date+"_log.txt")
If (Test-Path $TranscriptTarget) {Remove-Item $TranscriptTarget -Force}
Move-Item $Transcript $TranscriptTarget -Force
#################### Transcript Save
Write-host "Done (transcript saved)"