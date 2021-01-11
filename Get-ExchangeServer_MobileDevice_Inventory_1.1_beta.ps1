#########################################################################################
# LEGAL DISCLAIMER
# This Sample Code is provided for the purpose of illustration only and is not
# intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
# RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
# EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
# MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
# nonexclusive, royalty-free right to use and modify the Sample Code and to
# reproduce and distribute the object code form of the Sample Code, provided
# that You agree: (i) to not use Our name, logo, or trademarks to market Your
# software product in which the Sample Code is embedded; (ii) to include a valid
# copyright notice on Your software product in which the Sample Code is embedded;
# and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
# against any claims or lawsuits, including attorneys’ fees, that arise or result
# from the use or distribution of the Sample Code.
# 
# This posting is provided "AS IS" with no warranties, and confers no rights. Use
# of included script samples are subject to the terms specified at 
# https://www.microsoft.com/en-us/legal/intellectualproperty/copyright/default.aspx.
#
# Exchange Server Device partnership inventory
# Get-ExchangeServer_MobileDevice_Inventory_1.0.ps1
#  
#  Created by: Garrin Thompson 5/25/2020 garrint@microsoft.com
#  Updated by: Kevin Bloom 01/05/2021 Kevin.Bloom@Microsoft.com
#
#########################################################################################
# This script enumerates all devices in Exchange Server and reports on many properties of the
#   device/application and the mailbox owner.
#
# $deviceList is an array of hashtables, because deviceIDs may not be
#   unique in an environment. For instance when a device is configured with
#   two separate mailboxes in the same org, the same deviceID will appear twice.
#   Hashtables require uniqueness of the key so that's why the array of hashtable data 
#   structure was chosen.
#
# The devices can be sorted by a variety of properties like "LastPolicyUpdate" to determine 
#   stale partnerships or outdated devices needing to be removed.
# 
# The DisplayName of the user's CAS mailbox is recorded for importing with the 
#   Set-CasMailbox commandlet to configure allowedDeviceIDs. This is especially useful in 
#   scenarios where a migration to ABQ framework requires "grandfathering" in all or some
#   of the existing partnerships.
#
# Get-CasMailbox is run efficiently with the -HasActiveSyncDevicePartnership filter 
#
# Script is designed to be ran in Exchange Management Shell
#########################################################################################

# Writes output to a log file with a time date stamp
Function Write-Log {
	Param ([string]$string)
	$NonInteractive = 1
	# Get the current date
	[string]$date = Get-Date -Format G
	# Write everything to our log file
	( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
	# If NonInteractive true then supress host output
	if (!($NonInteractive)){
		( "[" + $date + "] - " + $string) | Write-Host
	}
}

# Set $OutputFolder to Current PowerShell Directory
write-progress -id 1 -activity "Beginning..." -PercentComplete (1) -Status "initializing variables"
[IO.Directory]::SetCurrentDirectory((Convert-Path (Get-Location -PSProvider FileSystem)))
$outputFolder = [IO.Directory]::GetCurrentDirectory()
$startDate = Get-Date
$logFile = $outputFolder + "\ExchangeServerMobileDevice_logfile_" + ($startDate).Ticks + ".txt"
$devicesOutput= $outputfolder + "\ExchangeServer Mobile Device Inventory_" + ($startDate).Ticks + ".csv"


# Clear the error log so that sending errors to file relate only to this run of the script
$error.clear()

# Get all mobiledevices from your tenant
write-progress -id 1 -Activity "Getting all Exchange Server Devices" -PercentComplete (5) -Status "Get-MobileDevice -ResultSize Unlimited"
$mobileDevices = Get-MobileDevice -ResultSize unlimited | Select-Object -Property friendlyname,deviceid,DeviceOS,DeviceModel,DeviceUseragent,devicetype,FirstSyncTime,WhenChangedUTC,identity,clientversion,clienttype,ismanaged,DeviceAccessState,DeviceAccessStateReason

# Measure the time the Invoke Command call takes to enumerate devices from Exchange Server
$progressActions = $mobileDevices.count
$invokeEndDate = Get-Date
$invokeElapsedTime = $invokeEndDate - $startDate
Write-Log ("Starting device collection");Write-Log ("Number of Devices found in Exchange Server:       " + ($progressActions));Write-Log ("Time to run Invoke command for Device retrieval:  " + ($($invokeElapsedTime)))
Write-host;write-host -foregroundcolor Magenta "Starting device collection";;sleep 2;write-host "-------------------------------------------------"
Write-Host -NoNewline "Number of Devices found in Exchange Server:       ";Write-Host -ForegroundColor Green $progressActions
Write-Host -NoNewline "Time to run Invoke command for Device retrieval:  ";write-host -ForegroundColor Yellow "$($invokeElapsedTime)"

# Get mailbox attributes for users with device partnerships from your tenant
Write-Progress -Id 1 -Activity "Getting all Exchange Server users with Devices" -PercentComplete (10) -Status "Get-CasMailbox -ResultSize Unlimited"
$mobileDeviceUsers = Get-CASMailbox -RecalculateHasActiveSyncDevicePartnership -ResultSize unlimited -Filter {HasActiveSyncDevicePartnership -eq "True"} | Select-Object -Property distinguishedname,displayname,id,primarysmtpaddress,activesyncmailboxpolicy,activesyncsuppressreadreceipt,activesyncdebuglogging,activesyncallowedids,activesyncblockeddeviceids

# Measure the time the get-casmailbox cmd takes to grab info for users with devices
$casMailboxUnlimitedEndDate = Get-Date
$casMailboxUnlimitedElapsedTime = $casMailboxUnlimitedEndDate - $invokeEndDate
Write-Log ("Number of Users with Devices in Exchange Server:  " + $($mobileDeviceUsers.count));Write-Log ("Time to run Get-CASMailbox -ResultSize Unlimited: " + $($casMailboxUnlimitedElapsedTime))
Write-Host -NoNewline "Number of Users with Devices in Exchange Server:  ";Write-Host -ForegroundColor Green "$($mobileDeviceUsers.count)"
Write-Host -NoNewline "Time to run Get-CASMailbox -ResultSize Unlimited: ";write-host -ForegroundColor Yellow "$($casMailboxUnlimitedElapsedTime)"

##
#  Now from the two arrays of hashtables, let's create a new array of hashtables containing calculated properties indexed by a property from the device list
#  This is a BIG LOOP!!

[System.Collections.ArrayList]$deviceList = New-Object System.Collections.ArrayList($null)
$currentProgress = 1
[TimeSpan]$caseCheckTotalTime=0

# Set some timedate variables for the stats report
$loopStartTime = Get-Date
$loopCurrentTime = Get-Date

foreach ($mobileDevice in $mobileDevices) {
    # The MobileDevice.Id has a consistent pattern in the directory, containing the mobile user's casmailbox id
      $userIndex = $mobileDevice.Identity.parent.parent.name
    Write-Progress -Id 1 -Activity "Getting all device partnerships from " -PercentComplete (5 + ($currentProgress/$progressActions * 90)) -Status "Enumerating a device for user $($userIndex)"
	#  UPDATE: In some cases, if a CASmailbox user ONLY has a REST partnership with Outlook for iOS / Android, the HasActiveSyncDevicePartnership will be false.
    if ((get-recipient -identity $userindex).RecipientTypeDetails -eq 'UserMailbox')
        {
        if($mobiledevice.ClientType.Value -eq "EAS")
            {
            # Powershell v4 allows super efficient handy reference of the array by an object value using the .where() method
            # I haven't tested this method with over 1000 users, so test here if efficiency results falter
            $mobileUser = $mobileDeviceUsers.where({$_.id -eq "$userIndex"})
            }
        Else 
            {
            $caseCheckStartDate = Get-Date
            if($userindex){
            $mobileUser = Get-CASMailbox -Identity $userIndex | Select-Object -Property distinguishedname,displayname,id,primarysmtpaddress,activesyncmailboxpolicy,activesyncsuppressreadreceipt,activesyncdebuglogging,activesyncallowedids,activesyncblockeddeviceids
            }
            else 
            {
                Write-Output "Could not find CASmailbox information for this device $mobileDevice" | Out-File $debugoutput -Append
                Write-Log ("Could not find CASmailbox information for this device $mobileDevice")
            }
            [timespan]$caseCheckEndTime = (Get-Date) - $caseCheckStartDate
            $caseCheckTotalTime += $caseCheckEndTime
            }
            
        # Setting the hashtable starting with the CASMailbox information
        # using shorthand notation for add-member
        $line = @{
            User=$userIndex
            DisplayName=$mobileUser.DisplayName
            PrimarySmtpAddress=$mobileUser.PrimarySmtpAddress
            Id=$mobileUser.Id
            ActivesyncSuppressReadReceipt=$mobileUser.activesyncsuppressreadreceipt
            ActivesyncDebugLogging=$mobileUser.activesyncdebuglogging
            DistinguishedName=$mobileUser.distinguishedname
            # Now including the MobileDevice information
            FriendlyName=$mobileDevice.friendlyname
            DeviceID=$mobileDevice.deviceid
            DeviceOS=$mobileDevice.DeviceOS
            DeviceModel=$mobileDevice.DeviceModel
            UserAgent=$mobileDevice.DeviceUserAgent
            FirstSyncTime=$mobileDevice.FirstSyncTime
            LastPolicyUpdate=$mobileDevice.WhenChangedUTC
            ClientProtocolVersion=$mobileDevice.clientversion
            ClientType=$mobileDevice.clienttype
            Managed=$mobileDevice.ismanaged
            AccessState=$mobileDevice.DeviceAccessState
            AccessReason=$mobileDevice.DeviceAccessStateReason
            # The RemovalID has some caveats before it can be re-used for the remove-mobiledevice commandlet. 
                # When exporting this value to a .csv file, there is a special character called, "section sign" § that gets 
                # converted to a '?' so we adjust for that with a regex in the "ABQ_remove.ps1" script example end of this script.
            RemovalID="$($mobiledevice.Identity.Parent)/$($mobiledevice.Identity.Name)"
            }
        # out-null since arraylist.add returns highest index so this is a way to ignore that value
        $deviceList.Add((New-Object PSobject -property $line)) | Out-Null
        }
    else 
        {
        Write-Log ("$userIndex Recipient Type is $((get-recipient -identity $userindex).RecipientTypeDetails) and is not being included in the report")
        }
	$currentProgress++
}
	 
Write-Log ("Time to re-run Get-CasMailbox for REST devices:   " + $($caseCheckTotalTime))
Write-Host -NoNewLine "Time to re-run Get-CasMailbox for REST devices:   ";write-host -ForegroundColor Yellow "$($caseCheckTotalTime)"

# Now to put all that info into a spreadsheet. 
write-progress -id 1 -activity "Creating spreadsheet" -PercentComplete (96) -Status "$outputFolder"

$deviceList | select DisplayName,user,id,PrimarySMTPAddress,FriendlyName,UserAgent,FirstSyncTime,LastPolicyUpdate,DeviceOS,ClientProtocolVersion,ClientType,DeviceModel,deviceid,AccessState,AccessReason,ActivesyncSuppressReadReceipt,ActivesyncDebugLogging,Managed,distinguishedname,RemovalID | export-csv -path $devicesoutput -notypeinformation -Append
  
# Capture any PS errors and output to a file
	$errfilename = $outputfolder + "\ExchangeServer Mobile Device Errorlog_" + ($startDate).Ticks + ".txt" 
	write-progress -id 1 -activity "Error logging" -PercentComplete (99) -Status "$errfilename"
	foreach ($err in $error) 
	{  
	    $logdata = $null 
	    $logdata = $err 
	    if ($logdata) 
	    { 
		out-file -filepath $errfilename -Inputobject $logData -Append 
	    } 
	}

write-progress -id 1 -activity "Complete" -PercentComplete (100) -Status "Success!"
$endDate = Get-Date
$elapsedTime = $endDate - $startDate
Write-Log ("Report started at    " + $($startDate));Write-Log ("Report ended at      " + $($endDate));Write-Log ("Total Elapsed Time:   " + $($elapsedTime)); Write-Log ("Device Collection Completed!")
write-host;Write-Host -NoNewLine "Report started at    ";write-host -ForegroundColor Yellow "$($startDate)"
Write-Host -NoNewLine "Report ended at      ";write-host -ForegroundColor Yellow "$($endDate)"
Write-Host -NoNewLine "Total Elapsed Time:   ";write-host -ForegroundColor Yellow "$($elapsedTime)"
Write-host "-------------------------------------------------";write-host -foregroundcolor Magenta "Device collection Complete! The logfile and the Exchange Server Mobile Device Inventory CSV was created in $outputFolder.";write-host;write-host;sleep 1
