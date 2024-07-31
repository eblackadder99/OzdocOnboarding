## Onboarding Script (365 Hybrid)
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

#region Powershell Admin check & set execution policy
## Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $scriptPath = $myinvocation.mycommand.definition
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
    $newProcess.Arguments = $arguments
    $newProcess.Verb = "runas";
    [System.Diagnostics.Process]::Start($newProcess);
    exit
}

Write-Host "Running as an administrator."
#endregion Powershell Admin Check

#region Check for updates

## Define the repository and file details
$owner = "eblackadder99"
$repo = "OzdocOnboarding"
$path = "Update/Version.txt"
$updateCheck = "C:\Ozdoc\UserOnboarding\Update\version.txt"

## Define the GitHub API URL
$url = "https://api.github.com/repos/$owner/$repo/contents/$path"

## Send a request to the GitHub API
$response = Invoke-RestMethod -Uri $url -Method Get -Headers @{"Accept"="application/vnd.github.v3+json"}
$base64Content = $response.content
$bytes = [System.Convert]::FromBase64String($base64Content)
$plaintext = [System.Text.Encoding]::UTF8.GetString($bytes).Trim()
$updateCheckContent = (Get-Content -Path $updateCheck -Raw).Trim()

Write-Host "Plaintext: $plaintext"
Write-Host "Update Check Content: $updateCheckContent"

if (-NOT($plaintext -eq $updateCheckContent)) {
    $updateAvailable = $true
}

if ($updateAvailable -eq $true) {
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File 'C:\Ozdoc\Update\UpdateOnboardingScript.ps1'"
}
#endregion Check for updates

#region PS module check

## List of modules to check
$modules = @("Microsoft.Graph.Users", "Microsoft.Graph.Groups", "ExchangeOnlineManagement", "DuoSecurity")

# Check if PowerShell 7 is running
$psVersion = $PSVersionTable.PSVersion.Major

if ($psVersion -ge 7) {
    Write-Output "PowerShell 7 or higher is running. Installing all modules."
} else {
    Write-Output "PowerShell version is less than 7. Installing all modules except DuoSecurity."
    $modules = $modules | Where-Object { $_ -ne "DuoSecurity" }
}

foreach ($module in $modules) {
    if (-not (Get-Module -Name $module -ListAvailable)) {
        Write-Output "$module is not installed. Installing now..."
        Install-Module -Name $module -Scope AllUsers -Force -AllowClobber
        $modulesInstalled = $true
    }
}

if ($modulesInstalled) {
    Write-Output "Required modules have been installed."
} else {
    Write-Output "All required modules were already installed."
}

## Check if each module is imported
foreach ($module in $modules) { 
    if (-not (Get-Module -Name $module -ListAvailable)) {
        Write-Output "$module is not imported. Importing now..."
        Import-Module -Name $module
    } else {
        Write-Output "$module is imported."
    }
}

#endregion PS module check

#region User Creation Form

## Create the form
$onboardingUserForm = New-Object System.Windows.Forms.Form
$onboardingUserForm.Text = 'New User Information'
$onboardingUserForm.Size = New-Object System.Drawing.Size(350,600)
$onboardingUserForm.StartPosition = 'CenterScreen'

## Decode the logo from Base64
$base64 = "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"
$bytes = [System.Convert]::FromBase64String($base64)
$ms = New-Object System.IO.MemoryStream($bytes, 0, $bytes.Length)
$ozdocLogoImage = [System.Drawing.Image]::FromStream($ms)
$ms.Seek(0, [System.IO.SeekOrigin]::Begin) | Out-Null

$iconSize = New-Object System.Drawing.Size(32,32)
$iconBitmap = New-Object System.Drawing.Bitmap($ozdocLogoImage, $iconSize)
$ozdocLogoIcon = [System.Drawing.Icon]::FromHandle($iconBitmap.GetHicon())

$onboardingUserForm.Icon = $ozdocLogoIcon

$ozdocLogo = New-Object System.Windows.Forms.PictureBox
$ozdocLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$ozdocLogo.Location = New-Object System.Drawing.Point(200,0)
$ozdocLogo.Width = 125
$ozdocLogo.Height = 30
$ozdocLogo.Image = $ozdocLogoImage
$onboardingUserForm.Controls.Add($ozdocLogo)

$ozdocLogo.Add_Click({
    Start-Process "https://github.com/eblackadder99/OzdocOnboarding"
})

$ms.Dispose()

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

## Get active email address domains from AD
$ADUsers = Get-ADUser -Filter * -Properties EmailAddress
$ADDomains = $ADUsers | ForEach-Object {
    if ($_.EmailAddress) {
        $_.EmailAddress.Split('@')[1]
    }
} | Sort-Object -Unique

## Create a textbox for new logon name
$newUserLogonName = New-Object System.Windows.Forms.TextBox
$newUserLogonName.Location = New-Object System.Drawing.Point(15,200)
$newUserLogonName.Size = New-Object System.Drawing.Size(135,23)
$onboardingUserForm.Controls.Add($newUserLogonName)
$newUserLogonName.Name = "newUserLogonNameField"
$newUserLogonName.Add_MouseEnter({
    $newUserLogonName.BackColor = [System.Drawing.Color]::LightCyan
})
$newUserLogonName.Add_MouseLeave({
    $newUserLogonName.BackColor = [System.Drawing.Color]::White
})

## Create a ComboBox for email domain selection
$domainCombobox = New-Object System.Windows.Forms.ComboBox
$domainCombobox.Location = New-Object System.Drawing.Point(170, 200)
$domainCombobox.Size = New-Object System.Drawing.Size(150, 20)
$ADDomains | ForEach-Object {
    $domainCombobox.Items.Add($_)
}
$onboardingUserForm.Controls.Add($domainCombobox)


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

## Adds a button for password generator
$generatePasswordButton = New-Object System.Windows.Forms.Button
$generatePasswordButton.Location = New-Object System.Drawing.Point(243,226)
$generatePasswordButton.Size = New-Object System.Drawing.Size(70,22)
$generatePasswordButton.Text = "Generate"
$onboardingUserForm.Controls.Add($generatePasswordButton)
$generatePasswordButton.Name = "generatePasswordButton"

## Event handler for button click
$generatePasswordButton.Add_Click({
    ## Generate the password
    $generatedPassword = Invoke-RestMethod -Uri "https://www.dinopass.com/password/strong" -Method Get
    $newPassword.Text = $generatedPassword
})

## Create a textbox for new account description
$newDescription = New-Object System.Windows.Forms.TextBox
$newDescription.Location = New-Object System.Drawing.Point(15,300)
$newDescription.Size = New-Object System.Drawing.Size (300,23)
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

## Event handler for mailbox selection check
$mailboxCheckedChanged = {
    if ($mailboxAccessCheck.Checked) {
        ## Shows the mailbox form if the checkbox selected
        $mailboxSelectionForm.ShowDialog()
    }
}

## Create the mailbox form
$mailboxSelectionForm = New-Object System.Windows.Forms.Form
$mailboxSelectionForm.Text = 'Mailbox Access Selection'
$mailboxSelectionForm.Size = New-Object System.Drawing.Size(450,350)
$mailboxSelectionForm.StartPosition = 'CenterScreen'

## Add the Ozdoc logo
$ozdocLogo = New-Object System.Windows.Forms.PictureBox
$ozdocLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$ozdocLogo.Location = New-Object System.Drawing.Point (200,0)
$ozdocLogo.Width = 125
$ozdocLogo.Height = 30
$ozdocLogoImage = $ozdocLogoImage
$ozdocLogo.Image = $OzdocLogoImage
$onboardingUserForm.Controls.Add($ozdocLogo)
$ozdocLogo.Add_Click({
    Start-Process "https://www.ozdoc.com.au"
})

## Create Enter button
$enterButton = New-Object System.Windows.Forms.Button
$enterButton.Location = New-Object System.Drawing.Point(125,275)
$enterButton.Size = New-Object System.Drawing.Size(75,23)
$enterButton.Text = 'Enter'
$enterButton.Add_MouseEnter({
    $enterButton.BackColor = [System.Drawing.Color]::LightBlue
})
$enterButton.Add_MouseLeave({
    $enterButton.BackColor = [System.Drawing.Color]::Transparent
})
$mailboxSelectionForm.AcceptButton = $enterButton
$mailboxSelectionForm.Controls.Add($enterButton)

$enterButton.Add_Click({
    $mailboxSelectionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $mailboxSelectionForm.Close()
})

## Creates cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(225,275)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$cancelButton.Add_MouseEnter({
    $cancelButton.BackColor = [System.Drawing.Color]::LightBlue
})
$cancelButton.Add_MouseLeave({
    $cancelButton.BackColor = [System.Drawing.Color]::Transparent
})
$mailboxSelectionForm.CancelButton = $cancelButton
$mailboxSelectionForm.Controls.Add($cancelButton)

## Header label
$mailboxHeaderLabel = New-Object System.Windows.Forms.Label
$mailboxHeaderLabel.Location = New-Object System.Drawing.Point(13,5)
$mailboxHeaderLabel.AutoSize = $true
$mailboxHeaderLabel.Text = 'Please add mailboxes to assign access to'
$mailboxSelectionForm.Controls.Add($mailboxHeaderLabel)

## Access level label
$accessSelectionLabel = New-Object System.Windows.Forms.Label
$accessSelectionLabel.Location = New-Object System.Drawing.Point(13,200)
$accessSelectionLabel.AutoSize = $true
$accessSelectionLabel.Text = 'What level of access is required?'
$mailboxSelectionForm.Controls.Add($accessSelectionLabel)

## Create the mailbox texbox
$mailboxTextbox = New-Object System.Windows.Forms.TextBox
$mailboxTextbox.Location = New-Object System.Drawing.Point(10,35)
$mailboxTextbox.Width = 412
$mailboxSelectionForm.Controls.Add($mailboxTextbox)
$mailboxSelectionForm.Add_Shown({
    $mailboxTextbox.Focus()
})

## Create the mailbox list box
$ListBox = New-Object System.Windows.Forms.ListBox
$ListBox.Location = New-Object System.Drawing.Point(10,98)
$ListBox.Size = New-Object System.Drawing.Size (412,105)
$listBox.SelectionMode = 'MultiExtended'
$mailboxSelectionForm.Controls.Add($ListBox)

## Creates the add button
$mailboxSelectionButton = New-Object System.Windows.Forms.Button
$mailboxSelectionButton.Location = New-Object System.Drawing.Point(10,63)
$mailboxSelectionButton.AutoSize = $true
$mailboxSelectionButton.Text = 'Add Mailbox'
$mailboxSelectionForm.Controls.Add($mailboxSelectionButton)

## Creates the remove button
$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Location = New-Object System.Drawing.Point(315, 63)
$removeButton.AutoSize = $true
$removeButton.Text = 'Remove Mailbox'
$mailboxSelectionForm.Controls.Add($removeButton)

## Event handler for the 'Remove' button
$removeButton_Click = {
    ## Remove the selected item from the list box
    if ($listBox.SelectedIndex -ne -1) {  # Check if an item is selected
        $listBox.Items.RemoveAt($listBox.SelectedIndex)
    }
}

## Adds a combo box to select what access type will need to be applied
$mailboxAccessType = New-Object System.Windows.Forms.ComboBox
$mailboxAccessType.Location = New-Object System.Drawing.Size(10,222)
$mailboxAccessType.Size = New-Object System.Drawing.Size (300,25)
$mailboxAccessType.Items.Add("Full Access")
$mailboxAccessType.Items.Add("Send As")
$mailboxAccessType.Items.Add("Read Only")
$mailboxSelectionForm.Controls.Add($mailboxAccessType)
$mailboxAccessType.Name = "mailboxAccessTypeField"

## Event handler for the button
$mailboxSelectionButton_Click = {
    $listBox.Items.Add($mailboxTextbox.Text)
    $mailboxTextbox.Clear()
    $mailboxTextbox.Focus()
}

## Register events
$mailboxSelectionButton.add_Click($mailboxSelectionButton_Click)
$removeButton.add_Click($removeButton_Click)

## Displays the form
$mailboxSelectionForm.Topmost = $true

foreach ($item in $mailbox) {
    Write-Output "You selected: $item"
}

## Register the event
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

## New AD account @ character label
$newUserLogonNameAtLabel = New-Object System.Windows.Forms.Label
$newUserLogonNameAtLabel.Location = New-Object System.Drawing.Point(152,203)
$newUserLogonNameAtLabel.AutoSize = $true
$newUserLogonNameAtLabel.Text = '@'
$onboardingUserForm.Controls.Add($newUserLogonNameAtLabel)

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

## Set required mailbox access
$mailboxAccessLabel = New-Object System.Windows.Forms.Label
$mailboxAccessLabel.Location = New-Object System.Drawing.Point(13,425)
$mailboxAccessLabel.AutoSize = $true
$mailboxAccessLabel.Text = 'Does the user require access to any mailboxes?'
$onboardingUserForm.Controls.Add($mailboxAccessLabel)

## Creates a LinkLabel for Sean's suggestions
$suggestionsLinkLabel = New-Object System.Windows.Forms.LinkLabel
$suggestionsLinkLabel.Text = "Have any suggestions?"
$suggestionsLinkLabel.Location = New-Object System.Drawing.Point(100,490)
$suggestionsLinkLabel.Autosize = $true

# Set the link area
$suggestionsLinkLabel.LinkArea = New-Object System.Windows.Forms.LinkArea(0,35)

# Add an event handler for the LinkClicked event
$suggestionsLinkLabel.add_LinkClicked({
    Start-Process "https://m.media-amazon.com/images/I/51RAulxObvS._AC_UF894,1000_QL80_.jpg"
})

# Add the LinkLabel to the form
$onboardingUserForm.Controls.Add($suggestionsLinkLabel)

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
$Required6Label.Location = New-Object System.Drawing.Point(($domainCombobox.Location.X + $domainCombobox.Width), $domainCombobox.Location.Y)
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
    $newUserLogonName = "$($newUserLogonName.Text)@$($domainCombobox.Text)"
    $newPassword = $newPassword.Text
    $newDescription = $newDescription.Text
    $365LicenseSelection = $365LicenseSelection.Text
    $ADSyncCheck = $ADSyncCheck.Checked
    $assignMailboxCheck = $mailboxAccessCheck.Checked
    $splitName = $newDisplayName -split ' '
    $newFirstName = $splitName[0]
    $newLastName = $splitName[1]
    $newName = "$newFirstName $newLastName"
    if ($assignMailboxCheck -eq $True) {
    $mailboxAccessType = $mailboxAccessType.Text
    $mailbox = $ListBox.Items | ForEach-Object { $_.ToString() }
    }
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

Write-Host `n "SAMAccountName = $newSAMAccountName`r"

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
    $mailboxAccess = "ReadPermission"
    Write-Host `n"Assigning read only access to selected mailboxes`r"
}

if ($assignMailboxCheck -eq $true) {
    Connect-ExchangeOnline
    foreach ($mbx in $mailbox) {
        try {
            Add-MailboxPermission -Identity $mbx -User $newUserLogonName -AccessRights $mailboxAccess
            Write-Host `n"Successfully assigned $mailboxAccess access to $mbx for $newUserLogonName`r"
            
            if ($mailboxAccessType -eq "Full Access") {
                Add-RecipientPermission -Identity $mbx -Trustee $newUserLogonName -AccessRights SendAs
                Write-Host `n"Successfully assigned SendAs permission to $mbx for $newUserLogonName`r"
            }
        } catch {
            Write-Host `n"Failed to assign $mailboxAccess access to $mbx for $newUserLogonName. Error: $_`r"
        }
    }
}

#endregion Assign mailbox Access

#region End Script

Write-Host `n "Account creation process complete, stopping script`r"
Read-Host -Prompt "Press Enter to continue"
exit

#endregion End Script