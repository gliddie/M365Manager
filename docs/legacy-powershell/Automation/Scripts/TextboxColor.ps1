###########################################
#
# Wayne Lindimore
# wlindimore@gmail.com
# AdminsCache.Wordpress.com
#
# 10-18-13
# PowerShell GUI Allowing Color Selection
##########################################
Add-Type -AssemblyName System.Windows.Forms

# Main Form 
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Font = "Comic Sans MS,10"
$mainForm.Text = " Color Dialog Demo"
$mainForm.ForeColor = "White"
$mainForm.BackColor = "DarkBlue"
$mainForm.Width = 450
$mainForm.Height = 330

# Alpha Label
$alphaLabel = New-Object System.Windows.Forms.Label
$alphaLabel.Location = "45,20"
$alphaLabel.Height = 22
$alphaLabel.Width = 100
$alphaLabel.Text = "Alpha"
$mainForm.Controls.Add($alphaLabel)
# Red Label
$redLabel = New-Object System.Windows.Forms.Label
$redLabel.Location = "45,50"
$redLabel.Height = 22
$redLabel.Width = 100
$redLabel.Text = "Red"
$mainForm.Controls.Add($redLabel)
# Green Label
$greenLabel = New-Object System.Windows.Forms.Label
$greenLabel.Location = "45,80"
$greenLabel.Height = 22
$greenLabel.Width = 100
$greenLabel.Text = "Green"
$mainForm.Controls.Add($greenLabel)
# Blue Label
$blueLabel = New-Object System.Windows.Forms.Label
$blueLabel.Location = "45,110"
$blueLabel.Height = 22
$blueLabel.Width = 100
$blueLabel.Text = "Blue"
$mainForm.Controls.Add($blueLabel)
# Color Name Label
$colorNameLabel = New-Object System.Windows.Forms.Label
$colorNameLabel.Location = "45,140"
$colorNameLabel.Height = 22
$colorNameLabel.Width = 120
$colorNameLabel.Text = "Color Name/Hex"
$mainForm.Controls.Add($colorNameLabel)

#Alpha TextBox
$alphaTextBox = New-Object System.Windows.Forms.TextBox
$alphaTextBox.Location = "180,20"
$alphaTextBox.Size = "100,20"
$alphaTextBox.ForeColor = "MediumBlue"
$alphaTextBox.BackColor = "White"
$mainForm.Controls.Add($alphaTextBox)
# Red TextBox
$redTextBox = New-Object System.Windows.Forms.TextBox
$redTextBox.Location = "180,50"
$redTextBox.Size = "100,20"
$redTextBox.ForeColor = "MediumBlue"
$redTextBox.BackColor = "White"
$mainForm.Controls.Add($redTextBox)
# Green TextBox
$greenTextBox = New-Object System.Windows.Forms.TextBox
$greenTextBox.Location = "180,80"
$greenTextBox.Size = "100,20"
$greenTextBox.ForeColor = "MediumBlue"
$greenTextBox.BackColor = "White"
$mainForm.Controls.Add($greenTextBox)
# Blue TextBox
$blueTextBox = New-Object System.Windows.Forms.TextBox
$blueTextBox.Location = "180,110"
$blueTextBox.Size = "100,20"
$blueTextBox.ForeColor = "MediumBlue"
$blueTextBox.BackColor = "White"
$mainForm.Controls.Add($blueTextBox)
# Color Name TextBox
$colorNameTextBox = New-Object System.Windows.Forms.TextBox
$colorNameTextBox.Location = "180,140"
$colorNameTextBox.Size = "100,20"
$colorNameTextBox.ForeColor = "MediumBlue"
$colorNameTextBox.BackColor = "White"
$mainForm.Controls.Add($colorNameTextBox)

# Color Demo TextBox
$colorDemoTextBox = New-Object System.Windows.Forms.TextBox
$colorDemoTextBox.Font = "Comic Sans MS,28"
$colorDemoTextBox.Location = "130,200"
$colorDemoTextBox.Size = "200,200"
$colorDemoTextBox.ForeColor = "MediumBlue"
$colorDemoTextBox.BackColor = "White"
$colorDemoTextBox.Text = "Color  1234"
$mainForm.Controls.Add($colorDemoTextBox)

# Lookup Button
$lookupButton = New-Object System.Windows.Forms.Button 
$lookupButton.Location = "300,20"
$lookupButton.Size = "85,28"
$lookupButton.ForeColor = "Green"
$lookupButton.BackColor = "White"
$lookupButton.Text = "Pick Color"
$lookupButton.add_Click({PickColor})
$mainForm.Controls.Add($lookupButton)

# Exit Button 
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = "300,50"
$exitButton.Size = "85,28"
$exitButton.ForeColor = "Red"
$exitButton.BackColor = "White"
$exitButton.Text = "Exit"
$exitButton.add_Click({$mainForm.close()})
$mainForm.Controls.Add($exitButton)

function PickColor {
    $colorDialog = new-object System.Windows.Forms.ColorDialog 
    $colorDialog.AllowFullOpen = $true
    $colorDialog.ShowDialog()
    $alphaTextBox.Text = $colordialog.color.A 
    $redTextBox.Text = $colordialog.color.R  
    $greenTextBox.Text = $colordialog.color.G  
    $blueTextBox.Text = $colordialog.color.B  
    $colorNameTextBox.Text = $colorDialog.Color.Name
    $error.clear()
    $colorDemoTextBox.ForeColor = [Convert]::ToInt32($colorDialog.Color.Name, 16)
    If($error.count -gt 0) {
        $colorDemoTextBox.ForeColor = $colorDialog.Color.Name
    }
}

[void] $mainForm.ShowDialog()