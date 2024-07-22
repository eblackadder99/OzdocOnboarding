# Onboarding Script (365 Hybrid)
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

#region Powershell Admin check
## Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
    $newProcess.Arguments = $arguments
    $newProcess.Verb = "runas";
    [System.Diagnostics.Process]::Start($newProcess);
    exit
}
Write-Host "Running as an administrator."
#endregion Powershell Admin Check

#region User Creation Form

## Create the form
$onboardingUserForm = New-Object System.Windows.Forms.Form
$onboardingUserForm.Text = 'New User Information'
$onboardingUserForm.Size = New-Object System.Drawing.Size(350,600)
$onboardingUserForm.StartPosition = 'CenterScreen'

## Add the Ozdoc logo
$ozdocLogo = New-Object System.Windows.Forms.PictureBox
$ozdocLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$ozdocLogo.Location = New-Object System.Drawing.Point (200,0)
$ozdocLogo.Width = 125
$ozdocLogo.Height = 30
$scriptDir = Split-Path $MyInvocation.MyCommand.Path
$logoFilePath = Join-Path $scriptDir 'OzdocLogoScript.png'
$ozdocLogoImage = [System.Drawing.Image]::FromFile($logoFilePath)
$ozdocLogo.Image = $OzdocLogoImage
$onboardingUserForm.Controls.Add($ozdocLogo)
$ozdocLogo.Add_Click({
    Start-Process "https://www.ozdoc.com.au"
})

## Create Enter button
$enterButton = New-Object System.Windows.Forms.Button
$enterButton.Location = New-Object System.Drawing.Point(75,515)
$enterButton.Size = New-Object System.Drawing.Size(75,23)
$enterButton.Text = 'Enter'
$enterButton.Add_MouseEnter({
    $enterButton.BackColor = [System.Drawing.Color]::LightBlue
})
$enterButton.Add_MouseLeave({
    $enterButton.BackColor = [System.Drawing.Color]::Transparent
})
$onboardingUserForm.AcceptButton = $enterButton
$onboardingUserForm.Controls.Add($enterButton)
$enterButton.Add_Click({
    ## Create an array to hold the names of empty textboxes
    $emptyFields = @()

    ## Check each control in the form
    foreach ($control in $onboardingUserForm.Controls) {
        if (-NOT( $control.Name -like "newDescriptionField")) {
        if ($control -is [System.Windows.Forms.TextBox]) {
            ## Checks if any fields are empty
            if ([string]::IsNullOrEmpty($control.Text)) {
                ## Add the name of the empty textbox to the array
                $emptyFields += $control.Name
            }
        }
    }
}
    ## Check if there are any empty fields
    if ($emptyFields.Count -gt 0) {
        ## Add the names of the empty fields into a single string
        $emptyFieldsList = $emptyFields -join ', '
        ## Show an error message with the names of the empty fields
        [System.Windows.Forms.MessageBox]::Show("Please fill the following fields: $emptyFieldsList", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } else {
        [System.Windows.Forms.MessageBox]::Show("All fields are filled", "Validation Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $onboardingUserForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $onboardingUserForm.Close()
    }
})
## Creates cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(175,515)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$cancelButton.Add_MouseEnter({
    $cancelButton.BackColor = [System.Drawing.Color]::LightBlue
})
$cancelButton.Add_MouseLeave({
    $cancelButton.BackColor = [System.Drawing.Color]::Transparent
})
$onboardingUserForm.CancelButton = $cancelButton
$onboardingUserForm.Controls.Add($cancelButton)

## Create a textbox for AD account username to copy
$Username = New-Object System.Windows.Forms.TextBox
$Username.Location = New-Object System.Drawing.Point(15,50)
$Username.Size = New-Object System.Drawing.Size(300,23)
$onboardingUserForm.Controls.Add($Username)
$Username.Name = "UsernameField"
$Username.Add_MouseEnter({
    $Username.BackColor = [System.Drawing.Color]::LightCyan
})
$Username.Add_MouseLeave({
    $Username.BackColor = [System.Drawing.Color]::White
})

## Create a textbox for new AD account username
$newSAMAccountName = New-Object System.Windows.Forms.TextBox
$newSAMAccountName.Location = New-Object System.Drawing.Point(15,100)
$newSAMAccountName.Size = New-Object System.Drawing.Size(300,23)
$onboardingUserForm.Controls.Add($newSAMAccountName)
$newSAMAccountName.Name = "newSAMAccountNameField"
$newSAMAccountName.Add_MouseEnter({
    $newSAMAccountName.BackColor = [System.Drawing.Color]::LightCyan
})
$newSAMAccountName.Add_MouseLeave({
    $newSAMAccountName.BackColor = [System.Drawing.Color]::White
})

## Create a textbox for new display name
$newDisplayName = New-Object System.Windows.Forms.TextBox
$newDisplayName.Location = New-Object System.Drawing.Point(15,150)
$newDisplayName.Size = New-Object System.Drawing.Size(300,23)
$onboardingUserForm.Controls.Add($newDisplayName)
$newDisplayName.Name = "newDisplayNameField"
$newDisplayName.Add_MouseEnter({
    $newDisplayName.BackColor = [System.Drawing.Color]::LightCyan
})
$newDisplayName.Add_MouseLeave({
    $newDisplayName.BackColor = [System.Drawing.Color]::White
})

## Create a textbox for new logon name
$newUserLogonName = New-Object System.Windows.Forms.TextBox
$newUserLogonName.Location = New-Object System.Drawing.Point(15,200)
$newUserLogonName.Size = New-Object System.Drawing.Size(300,23)
$onboardingUserForm.Controls.Add($newUserLogonName)
$newUserLogonName.Name = "newUserLogonNameField"
$newUserLogonName.Add_MouseEnter({
    $newUserLogonName.BackColor = [System.Drawing.Color]::LightCyan
})
$newUserLogonName.Add_MouseLeave({
    $newUserLogonName.BackColor = [System.Drawing.Color]::White
})

## Create a textbox for new password
$newPassword = New-Object System.Windows.Forms.TextBox
$newPassword.Location = New-Object System.Drawing.Point(15,250)
$newPassword.Size = New-Object System.Drawing.Size(300,23)
$onboardingUserForm.Controls.Add($newPassword)
$newPassword.Name = "newPasswordField"
$newPassword.Add_MouseEnter({
    $newPassword.BackColor = [System.Drawing.Color]::LightCyan
})
$newPassword.Add_MouseLeave({
    $newPassword.BackColor = [System.Drawing.Color]::White
})

## Create a textbox for new account description
$newDescription = New-Object System.Windows.Forms.TextBox
$newDescription.Location = New-Object System.Drawing.Point(15,300)
$newDescription.Size = New-Object System.Drawing.Size(300,23)
$onboardingUserForm.Controls.Add($newDescription)
$newDescription.Name = "newDescriptionField"
$newDescription.Add_MouseEnter({
    $newDescription.BackColor = [System.Drawing.Color]::LightCyan
})
$newDescription.Add_MouseLeave({
    $newDescription.BackColor = [System.Drawing.Color]::White
})

## Adds a checkbox for AD Sync
$ADSyncCheck = New-Object System.Windows.Forms.CheckBox
$ADSyncCheck.Location = New-Object System.Drawing.Size(220,327)
$ADSyncCheck.Size = New-Object System.Drawing.Size (150,25)
$onboardingUserForm.Controls.Add($ADSyncCheck)
$ADSyncCheck.Name = "ADSyncCheckField"

## Adds a combo box to select what license will need to be applied
$365LicenseSelection = New-Object System.Windows.Forms.ComboBox
$365LicenseSelection.Location = New-Object System.Drawing.Size(15,370)
$365LicenseSelection.Size = New-Object System.Drawing.Size (300,25)
$365LicenseSelection.Items.Add("Business Premium")
$365LicenseSelection.Items.Add("Business Standard")
$365LicenseSelection.Items.Add("None")
$onboardingUserForm.Controls.Add($365LicenseSelection)
$365LicenseSelection.Name = "365LicenseSelectionField"

## Adds a checkbox for mailbox selection
$mailboxAccessCheck = New-Object System.Windows.Forms.CheckBox
$mailboxAccessCheck.Location = New-Object System.Drawing.Size(270,424)
$mailboxAccessCheck.Size = New-Object System.Drawing.Size (150,25)
$onboardingUserForm.Controls.Add($mailboxAccessCheck)
$mailboxAccessCheck.Name = "mailboxAccessCheckField"

## Event handler fpr mailbox selection check
$mailboxCheckedChanged = {
    if ($mailboxAccessCheck.Checked) {
        ## Shows the mailbox form if the checkbox selected
        $mailboxSelectionForm.ShowDialog()
    }
}
# Register the event
$mailboxAccessCheck.add_CheckedChanged($mailboxCheckedChanged)

## Header label
$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Location = New-Object System.Drawing.Point(13,5)
$headerLabel.AutoSize = $true
$headerLabel.Text = 'Please enter new user information'
$onboardingUserForm.Controls.Add($headerLabel)

## AD account to copy label
$UsernameLabel = New-Object System.Windows.Forms.Label
$UsernameLabel.Location = New-Object System.Drawing.Point(13,30)
$UsernameLabel.AutoSize = $true
$UsernameLabel.Text = 'Enter SAM username to copy - e.g. Example.User'
$onboardingUserForm.Controls.Add($UsernameLabel)

## New AD account username label
$newSAMAccountNameLabel = New-Object System.Windows.Forms.Label
$newSAMAccountNameLabel.Location = New-Object System.Drawing.Point(13,80)
$newSAMAccountNameLabel.AutoSize = $true
$newSAMAccountNameLabel.Text = 'Enter new SAM account name - e.g. Example.User'
$onboardingUserForm.Controls.Add($newSAMAccountNameLabel)

## New AD account displayname label
$newDisplayNameLabel = New-Object System.Windows.Forms.Label
$newDisplayNameLabel.Location = New-Object System.Drawing.Point(13,130)
$newDisplayNameLabel.AutoSize = $true
$newDisplayNameLabel.Text = 'Enter display name'
$onboardingUserForm.Controls.Add($newDisplayNameLabel)

## New AD account logon name label
$newUserLogonNameLabel = New-Object System.Windows.Forms.Label
$newUserLogonNameLabel.Location = New-Object System.Drawing.Point(13,180)
$newUserLogonNameLabel.AutoSize = $true
$newUserLogonNameLabel.Text = 'Enter email address - e.g. user@domain.com'
$onboardingUserForm.Controls.Add($newUserLogonNameLabel)

## New AD account password label
$newPasswordLabel = New-Object System.Windows.Forms.Label
$newPasswordLabel.Location = New-Object System.Drawing.Point(13,230)
$newPasswordLabel.AutoSize = $true
$newPasswordLabel.Text = 'Enter password'
$onboardingUserForm.Controls.Add($newPasswordLabel)

## New AD account description label
$newDescriptionLabel = New-Object System.Windows.Forms.Label
$newDescriptionLabel.Location = New-Object System.Drawing.Point(13,280)
$newDescriptionLabel.AutoSize = $true
$newDescriptionLabel.Text = 'Enter a description for the account'
$onboardingUserForm.Controls.Add($newDescriptionLabel)

## Check if AD sync needs to run
$ADSyncCheckLabel = New-Object System.Windows.Forms.Label
$ADSyncCheckLabel.Location = New-Object System.Drawing.Point(13,328)
$ADSyncCheckLabel.AutoSize = $true
$ADSyncCheckLabel.Text = 'Does the user require a 365 account?'
$onboardingUserForm.Controls.Add($ADSyncCheckLabel)

## Select license to apply label
$365LicenseSelectionLabel = New-Object System.Windows.Forms.Label
$365LicenseSelectionLabel.Location = New-Object System.Drawing.Point(13,350)
$365LicenseSelectionLabel.AutoSize = $true
$365LicenseSelectionLabel.Text = 'Which license needs to be applied?'
$onboardingUserForm.Controls.Add($365LicenseSelectionLabel)

## User site verification check
$siteVerificationLabel = New-Object System.Windows.Forms.Label
$siteVerificationLabel.Location = New-Object System.Drawing.Point(13,400)
$siteVerificationLabel.AutoSize = $true
$siteVerificationLabel.Text = 'Have you verified which site the user will be at?'
$onboardingUserForm.Controls.Add($siteVerificationLabel)

## 
$mailboxAccessLabel = New-Object System.Windows.Forms.Label
$mailboxAccessLabel.Location = New-Object System.Drawing.Point(13,425)
$mailboxAccessLabel.AutoSize = $true
$mailboxAccessLabel.Text = 'Does the user require access to any mailboxes?'
$onboardingUserForm.Controls.Add($mailboxAccessLabel)

#region Required Fields

## Create a red asterisk for each required textbox
$Required1Label = New-Object System.Windows.Forms.Label
$Required1Label.Text = "*"
$Required1Label.ForeColor = 'Red'
$Required1Label.Location = New-Object System.Drawing.Point(($Username.Location.X + $Username.Width), $Username.Location.Y)
$onboardingUserForm.Controls.Add($Required1Label)

$Required2Label = New-Object System.Windows.Forms.Label
$Required2Label.Text = "*"
$Required2Label.ForeColor = 'Red'
$Required2Label.Location = New-Object System.Drawing.Point(($newSAMAccountName.Location.X + $newSAMAccountName.Width), $newSAMAccountName.Location.Y)
$onboardingUserForm.Controls.Add($Required2Label)

$Required3Label = New-Object System.Windows.Forms.Label
$Required3Label.Text = "*"
$Required3Label.ForeColor = 'Red'
$Required3Label.Location = New-Object System.Drawing.Point(($newDisplayName.Location.X + $newDisplayName.Width), $newDisplayName.Location.Y)
$onboardingUserForm.Controls.Add($Required3Label)

$Required6Label = New-Object System.Windows.Forms.Label
$Required6Label.Text = "*"
$Required6Label.ForeColor = 'Red'
$Required6Label.Location = New-Object System.Drawing.Point(($newUserLogonName.Location.X + $newUserLogonName.Width), $newUserLogonName.Location.Y)
$onboardingUserForm.Controls.Add($Required6Label)

$Required7Label = New-Object System.Windows.Forms.Label
$Required7Label.Text = "*"
$Required7Label.ForeColor = 'Red'
$Required7Label.Location = New-Object System.Drawing.Point(($newPassword.Location.X + $newPassword.Width), $newPassword.Location.Y)
$onboardingUserForm.Controls.Add($Required7Label)

#endregion Asterisk

$onboardingUserForm.Topmost = $true

$onboardingUserForm.Add_Shown({ $Username.Select() })
$result = $onboardingUserForm.ShowDialog()

if ($result = $true)
{
    $Username = $Username.Text
    $newSAMAccountName = $newSAMAccountName.Text
    $newDisplayName = $newDisplayName.Text
    $newUserLogonName = $newUserLogonName.Text
    $newPassword = $newPassword.Text
    $newDescription = $newDescription.Text
    $365LicenseSelection = $365LicenseSelection.Text
    $ADSyncCheck = $ADSyncCheck.Checked
    $assignMailboxCheck = $mailboxAccessCheck.Checked
    $splitName = $newDisplayName -split ' '
    $newFirstName = $splitName[0]
    $newLastName = $splitName[1]
    $newName = "$newFirstName $newLastName"
    $mailboxAccessType = $mailboxAccessType.Text
    $mailbox = $ListBox.Items | ForEach-Object { $_.ToString() }

    Write-Host "All information has been entered"
}
#endregion User Creation Form

#region AD User Creation

## Get OU of user that is being copied
$new_OU_DN = (Get-ADUser $username -Properties distinguishedName).distinguishedName
$new_OU_DN = ($new_OU_DN -split ",",2)[1]

## Password config
$enableUserAfterCreation = $true
$passwordNeverExpires = $True
$cannotChangePassword = $false

## Params Attributes to copy from ADUser
$username = Get-Aduser $username -Properties memberOf, manager, title, department, company, streetAddress, City, POBox, State, PostalCode, Country, telephoneNumber, wWWHomePage, physicalDeliveryOfficeName

$params = @{'SamAccountName' = $newSAMAccountName;
            'Instance' = $Username;
            'DisplayName' = $newDisplayName;
            'GivenName' = $newFirstName;
            'SurName' = $newLastName;
            'PasswordNeverExpires' = $passwordNeverExpires;
            'CannotChangePassword' = $cannotChangePassword;
            'Description' = $newDescription;
            'Enabled' = $enableUserAfterCreation;
            'UserPrincipalName' = $newUserLogonName;
            'AccountPassword' = (ConvertTo-SecureString -AsPlainText $newPassword -Force);
        }

## Create the new user account
New-ADUser -Name $newName @params

## Mirror all the groups the original account was a member of
$username.Memberof | ForEach-Object {Add-ADGroupMember $_ $newSAMAccountName }

## Move the new user account into the assigned OU
Get-ADUser $newSAMAccountName| Move-ADObject -TargetPath $new_OU_DN

## Check if BSN-Employees exists then add the user to the group
$BSNExists = $false
$BSNGroup = Get-ADGroup -Filter { Name -eq "BSN-Employees" }
if ($null -ne $group) {
    $BSNExists = $true
    Add-ADGroupMember -Identity $BSNGroup -Members $newUserLogonName
}   else {
    Write-Host `n "BSN-Employees group does not exist in AD.`r"
}

## Display new user details
Write-Host `n "Displaying Settings for Account Created:`r"

Write-Host `n "sAMAccountName = $newSAMAccountName`r"

Write-Host `n "DisplayName = $newDisplayName`r"

Write-Host `n "FistName = $newFirstName`r"

Write-Host `n "LastName = $newLastName`r"

Write-Host `n "User LogonName = $newUserLogonName`r"

Write-Host `n "OU Path = $new_OU_DN`r"

Start-Sleep -Seconds 5

## Run AD Sync if required and check that user has synced before continuing

## Runs the ADSyncSyncCycle
if ($ADSyncCheck -eq "True") {
    Write-Host `n "Connecting to MS Graph service...`r"
    Connect-MgGraph -Scopes User.ReadWrite.All, Group.ReadWrite.All, Organization.Read.All -NoWelcome
    Write-Host `n "Running ADSync Service...`r"
    Start-ADSyncSyncCycle -PolicyType Delta
    Write-Host `n "Waiting for user to sync to 365...`r"
    }
    else {
        Write-Host `n "AD sync is not required. Ending script...`r"
    exit
}
## Check if the user exists
$userExists = $false

## Get the current time
$startTime = Get-Date

do {
    try {
        ## Attempt to get the user
        $syncedUser = Get-MgUser -UserId $newUserLogonName 2>$null

        ## If the user is found, set $userExists to $true
        if ($null -ne $syncedUser) {
            $userExists = $true
        }
    }
    catch {
        ## If an error occurs (e.g., the user is not found), wait for a while before trying again
        Start-Sleep -Seconds 10
    }

    ## Get the current time
    $currentTime = Get-Date
    $timeOut = 10
    ## Calculate the time difference
    $timeDifference = $currentTime - $startTime

    ## If more than 10 minutes has passed, break the loop
    if ($timeDifference.TotalMinutes -gt $timeOut) {
        Write-Host "Timeout: User was not found within $timeOut minutes."
        break
    }
} while (-not $userExists)
#endregion AD User Creation

#region 365 User Setup

## Assign UsageLocation to User
Write-Host `n "Assigning license to user`r"
Update-MgUser -UserID $newUserLogonName -Usagelocation 'AU'
$licenseSelected = $true

if ($365LicenseSelection -eq "Business Premium") {
    $365License = "SPB"
    Write-Host `n"Assigning Business Premium License`r"
}
elseif ($365LicenseSelection -eq "Business Standard") {
    $365License = "O365_BUSINESS_PREMIUM"
    Write-Host `n"Assigning Business Standard License`r"
}
elseif ($365LicenseSelection -eq "None") {
    Write-Host `n"No license has been selected, stopping script`r"
    $licenseSelected = $false
    exit
}

if ($licenseSelected -eq $true) {
    Write-Host `n "Applying 365 License`r"

    ## Get all SKUs
    $allSKUs = Get-MgSubscribedSku -all 2>$null

    ## Filter for Business Premium licenses
    if ($null -eq $365License) {
        Write-Host `n "No license found with ID '$365License', stopping script`r"
        exit
    }

    ## Select Business Premium license and assign to user
    $License = $allSKUs | Where-Object { $_.SkuPartNumber -eq "$365License" }
    Set-MgUserLicense -UserId $newUserLogonName -AddLicenses @{SkuId = $License.SkuId} -RemoveLicenses @()
    Write-Host `n "License $accountSkuID has been assigned to $newDisplayName`r"
}
if ($BSNExists -eq $false) {
## Below line may not be required?
$allGroups = Get-MgGroup
$365Group = Get-MgGroup -Filter "displayName eq 'BSN-Employees'"
$365User = Get-MgUser -UserId $newUserLogonName
if ($null -ne $365Group) {
    New-MgGroupMember -GroupId $365Group.Id -DirectoryObjectId $365User.Id
    Write-Host "User '$newUserLogonName' has been added to the group BSN-Employees."
} else {
    Write-Host "The BSN-Employees group does not exist in 365."
    }
}

#endregion 365 User Setup

#region Assign mailbox access

if ($mailboxAccessType -eq "Full Access") {
    $mailboxAccess = "FullAccess"
    Write-Host `n"Assigning full access to selected mailboxes`r"
}
elseif ($mailboxAccessType -eq "Send As") {
    $mailboxAccess = "SendAs"
    Write-Host `n"Assigning send as access to selected mailboxes`r"
}
elseif ($mailboxAccessType -eq "Read Only") {
    $mailboxAccess = "ReadOnly"
    Write-Host `n"Assigning read only access to selected mailboxes`r"
}

if ($assignMailboxCheck = $true) {
Connect-ExchangeOnline
foreach ($mbx in $mailbox) {
    Add-MailboxPermission -Identity $mbx -User $newUserLogonName -AccessRights $mailboxAccess
    }
}

#endregion Assign mailbox Access

#region End Script

Write-Host `n "Account creation process complete, stopping script`r"
Read-Host -Prompt "Press Enter to continue"
exit

#endregion End Script