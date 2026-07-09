# Create new Site in Active Directory
Function Input-Form
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "New 3 Letter Site" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 320 ; $form.Height = 220   # Make the form wider

    $Top = 10
    $Left = 120
  
    ## Regions
    $Script:lblRegion = New-Object System.Windows.Forms.Label   
        $Script:lblRegion.Text = "Select Region:"
        $Script:lblRegion.Top = $Top; $Script:lblRegion.Left = $Left-30; $Script:lblRegion.Width=150 ;$Script:lblRegion.AutoSize = $true
        $form.Controls.Add($Script:lblRegion)    # Add to Form 

    $Top = $Top + 20
    $Script:chkAsia = New-Object System.Windows.Forms.RadioButton
        $Script:chkAsia.Text = "Asia" 
        $Script:chkAsia.Top = $Top ; $Script:chkAsia.Left = $Left; $Script:chkAsia.Width=150 ;$Script:chkAsia.AutoSize = $true 
        $form.Controls.Add($Script:chkAsia)    # Add to Form

    $Top = $Top + 20
    $Script:chkEULA = New-Object System.Windows.Forms.RadioButton
        $Script:chkEULA.Text = "EULA" 
        $Script:chkEULA.Top = $Top ; $Script:chkEULA.Left = $Left; $Script:chkEULA.Width=150 ;$Script:chkEULA.AutoSize = $true 
        $form.Controls.Add($Script:chkEULA)    # Add to Form

    $Top = $Top + 20
    $Script:chkNA = New-Object System.Windows.Forms.RadioButton
        $Script:chkNA.Text = "NA" 
        $Script:chkNA.Top = $Top ; $Script:chkNA.Left = $Left; $Script:chkNA.Width=150 ;$Script:chkNA.AutoSize = $true 
        $form.Controls.Add($Script:chkNA)    # Add to Form

    $Top = $Top + 30
    ## 3 Letter Site Code
    $Script:lblSiteCode = New-Object System.Windows.Forms.Label   
        $Script:lblSiteCode.Text = "3 Letter Site Code:"
        $Script:lblSiteCode.Top = $Top; $Script:lblSiteCode.Left = $Left-50; $Script:lblSiteCode.Width=100
        $form.Controls.Add($Script:lblSiteCode)    # Add to Form 
        # 
        $Script:txtSiteCode = New-Object Windows.Forms.TextBox
        $Script:txtSiteCode.Top = $Top; $Script:txtSiteCode.Left = $Left+50; $Script:txtSiteCode.Width = 70;
        $Script:txtSiteCode.Text = ""
        $form.Controls.Add($Script:txtSiteCode)
}

Input-Form
Add-FormStandardButtons
Publish-Form

If ($Global:Result -eq "OK")
{   
    $3LtrLoc = ($Script:txtSiteCode.Text).ToUpper()
    If ($Script:chkAsia.Checked -eq $True)
    {
        $ADPath = "OU=ASIA,DC=global,DC=ul,DC=com"
    }
    else
    {
        If ($Script:chkEULA.Checked -eq $True)
        {
            $ADPath = "OU=EULA,DC=global,DC=ul,DC=com"
        }
        else
        {
            If ($Script:chkNA.Checked -eq $True)
            {
                $ADPath = "OU=NA,DC=global,DC=ul,DC=com"
            }
        }                
    }

    #Create the top level OU for this site

    $Name = $3LtrLoc + "Win7"
    write-host "Creating the Site: " $Name -ForegroundColor Green
    New-ADOrganizationalUnit -Name $Name -Path $ADPath

    #Create the subOU's for this site
    $ADPath = "OU=" + $Name + "," + $ADPath
    write-host "Creating the Computers OU in: " $ADPath -ForegroundColor Green
    New-ADOrganizationalUnit -Name "Computers" -Path $ADPath
    write-host "Creating the LabPCs OU in: " $ADPath -ForegroundColor Green
    New-ADOrganizationalUnit -Name "LabPCs" -Path $ADPath
    write-host "Creating the Users OU in: " $ADPath -ForegroundColor Green
    New-ADOrganizationalUnit -Name "Users" -Path $ADPath
}
else
{
    write-host "Creation of new 3 letter site cancelled" -ForegroundColor Red
}