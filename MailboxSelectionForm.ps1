## Load required assemblies
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

## Create the form
$mailboxSelectionForm = New-Object System.Windows.Forms.Form
$mailboxSelectionForm.Text = 'Mailbox Access Selection'
$mailboxSelectionForm.Size = New-Object System.Drawing.Size(450,350)
$mailboxSelectionForm.StartPosition = 'CenterScreen'

## Add the Ozdoc logo
$ozdocLogo = New-Object System.Windows.Forms.PictureBox
$ozdocLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$ozdocLogo.Location = New-Object System.Drawing.Point (310,2)
$ozdocLogo.Width = 125
$ozdocLogo.Height = 30
$scriptDir = Split-Path $MyInvocation.MyCommand.Path
$logoFilePath = Join-Path $scriptDir 'OzdocLogoScript.png'
$ozdocLogoImage = [System.Drawing.Image]::FromFile($logoFilePath)
$ozdocLogo.Image = $OzdocLogoImage
$mailboxSelectionForm.Controls.Add($ozdocLogo)
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

## 
$accessSelectionLabel = New-Object System.Windows.Forms.Label
$accessSelectionLabel.Location = New-Object System.Drawing.Point(13,200)
$accessSelectionLabel.AutoSize = $true
$accessSelectionLabel.Text = 'What level of access is required?'
$mailboxSelectionForm.Controls.Add($accessSelectionLabel)

## Create the text box
$mailboxTextbox = New-Object System.Windows.Forms.TextBox
$mailboxTextbox.Location = New-Object System.Drawing.Point(10,35)
$mailboxTextbox.Width = 412
$mailboxSelectionForm.Controls.Add($mailboxTextbox)

## Create the list box
$ListBox = New-Object System.Windows.Forms.ListBox
$ListBox.Location = New-Object System.Drawing.Point(10,98)
$ListBox.Size = New-Object System.Drawing.Size (412,105)
$listBox.SelectionMode = 'MultiExtended'
$mailboxSelectionForm.Controls.Add($ListBox)

## Create add the button
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

# Event handler for the 'Remove' button
$removeButton_Click = {
    # Remove the selected item from the list box
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
$mailboxSelectionForm.ShowDialog()

foreach ($item in $mailbox) {
    Write-Output "You selected: $item"
}