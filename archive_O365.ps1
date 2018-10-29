##############################################################
#  Script     : archive_O365.ps1
#  Author     : Trent Anderson
#  Date       : 20180927
#  Last Edited: 20181029, Trent Anderson
#  Description: Uses Veeam O365 to export users data
#  		Must be run inside a 
#  		Veeam Backup for Microsoft Office 365 PowerShell session
##############################################################

# Get archive folder path
Function Get-Folder($initialDirectory)

{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select an archive location"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}

$archive_destination = Get-Folder
if (([string]::IsNullOrEmpty($archive_destination)))
{
	exit
}

# Search AD for user accounts that are disabled and return username
$users = Search-ADAccount -AccountDisabled -UsersOnly | findstr SamAccountName |  foreach-object {$_.split(" ")[9]}

# Initiate Veeam restore session for email
$session = Start-VBOExchangeItemRestoreSession -LatestState
$database = Get-VEXDatabase -session $session

foreach ($user in $users)
{
        $fullname = get-aduser -identity $user -properties name | select-object -expandproperty name
        $mailbox = Get-VEXMailbox -database $database -name $fullname

	# Check if a mailbox exists for a given user
        if (([string]::IsNullOrEmpty($mailbox)))
        {
                write-host "$fullname doesn't have a mailbox"
        }

	# Export mailbox
        else
        {
                write-host "Exporting mailbox for $fullname"
                Export-VEXItem -Mailbox $mailbox -to $archive_destination\${user}-mailbox.pst
        }

}

# Close Veeam restore session
Stop-VBOExchangeItemRestoreSession -Session $session
write-host "VEOExchangeItemRestore Session stopped"

# Initiate Veeam save session
$job = Get-VBOJob -Name "General Backup"
Start-VEODRestoreSession -Job $job -Server localhost
$session = Get-VEODRestoreSession

foreach ($user in $users)
{
    $fullname = get-aduser -identity $user -properties name | select-object -expandproperty name
	$oduser =  Get-VEODUser -Session $session -Name $fullname

    # Check if a OneDrive folder exists for a given user
    if (([string]::IsNullOrEmpty($oduser)))
    {
       write-host "$fullname doesn't have OneDrive"
    }

    # Export OneDrive
    else
    {
        write-host "Exporting OneDrive for $fullname"
        Save-VEODDocument -User $oduser -Path "$archive_destination\${user}-onedrive.zip" -AsZip
    }

}

# Close Veeam restore session
Stop-VEODRestoreSession -Session $session
write-host "VEODRestore Session stopped"
