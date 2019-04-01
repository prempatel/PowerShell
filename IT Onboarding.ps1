<#
.NAME
    New Employee Onboarding
    #requires -modules ActiveDirectory
#>

Import-Module ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$NewEmployeeForm = New-Object system.Windows.Forms.Form
$NewEmployeeForm.ClientSize = '800,290'
$NewEmployeeForm.text = "IT Onboarding"
$NewEmployeeForm.StartPosition = "CenterScreen"
#$NewEmployeeForm.FormBorderStyle = 'Fixed3D'  //Disables resizing of window
$NewEmployeeForm.MaximizeBox = $false

$UserInfoGroup = New-Object system.Windows.Forms.Groupbox
$UserInfoGroup.height = 171
$UserInfoGroup.width = 400
$UserInfoGroup.location = New-Object System.Drawing.Point(10, 25)

$ITOnboardLabel = New-Object system.Windows.Forms.Label
$ITOnboardLabel.text = "IT Onboarding"
$ITOnboardLabel.AutoSize = $true
$ITOnboardLabel.width = 25
$ITOnboardLabel.height = 10
$ITOnboardLabel.location = New-Object System.Drawing.Point(145, 5)
$ITOnboardLabel.Font = 'Microsoft Sans Serif,12,style=Bold'

$FirstNameLabel = New-Object system.Windows.Forms.Label
$FirstNameLabel.text = "First Name"
$FirstNameLabel.AutoSize = $true
$FirstNameLabel.width = 25
$FirstNameLabel.height = 10
$FirstNameLabel.location = New-Object System.Drawing.Point(66, 15)
$FirstNameLabel.Font = 'Microsoft Sans Serif,10'

$FirstNameTextBox = New-Object system.Windows.Forms.TextBox
$FirstNameTextBox.multiline = $false
$FirstNameTextBox.width = 120
$FirstNameTextBox.height = 20
$FirstNameTextBox.location = New-Object System.Drawing.Point(138, 15)
$FirstNameTextBox.Font = 'Microsoft Sans Serif,10'

$LastNameLabel = New-Object system.Windows.Forms.Label
$LastNameLabel.text = "Last Name"
$LastNameLabel.AutoSize = $true
$LastNameLabel.width = 25
$LastNameLabel.height = 10
$LastNameLabel.location = New-Object System.Drawing.Point(67, 40)
$LastNameLabel.Font = 'Microsoft Sans Serif,10'

$LastNameTextBox = New-Object system.Windows.Forms.TextBox
$LastNameTextBox.multiline = $false
$LastNameTextBox.width = 120
$LastNameTextBox.height = 20
$LastNameTextBox.location = New-Object System.Drawing.Point(138, 40)
$LastNameTextBox.Font = 'Microsoft Sans Serif,10'

$EmailLabel = New-Object system.Windows.Forms.Label
$EmailLabel.text = "Email Address"
$EmailLabel.AutoSize = $true
$EmailLabel.width = 25
$EmailLabel.height = 10
$EmailLabel.location = New-Object System.Drawing.Point(45, 68)
$EmailLabel.Font = 'Microsoft Sans Serif,10'

$EmailTextBox = New-Object system.Windows.Forms.TextBox
$EmailTextBox.multiline = $false
$EmailTextBox.width = 120
$EmailTextBox.height = 20
$EmailTextBox.location = New-Object System.Drawing.Point(138, 65)
$EmailTextBox.Font = 'Microsoft Sans Serif,10'

$EmailSuffixLabel = New-Object system.Windows.Forms.Label
$EmailSuffixLabel.text = "@takarabelmont.com"
$EmailSuffixLabel.AutoSize = $false
$EmailSuffixLabel.width = 136
$EmailSuffixLabel.height = 20
$EmailSuffixLabel.location = New-Object System.Drawing.Point(260, 69)
$EmailSuffixLabel.Font = 'Microsoft Sans Serif,10'

$PasswordLabel = New-Object system.Windows.Forms.Label
$PasswordLabel.text = "Password"
$PasswordLabel.AutoSize = $true
$PasswordLabel.width = 25
$PasswordLabel.height = 10
$PasswordLabel.location = New-Object System.Drawing.Point(73, 90)
$PasswordLabel.Font = 'Microsoft Sans Serif,10'

$PasswordMaskedTextBox = New-Object system.Windows.Forms.MaskedTextBox
$PasswordMaskedTextBox.PasswordChar = '*'
$PasswordMaskedTextBox.multiline = $false
$PasswordMaskedTextBox.width = 120
$PasswordMaskedTextBox.height = 20
$PasswordMaskedTextBox.location = New-Object System.Drawing.Point(138, 90)
$PasswordMaskedTextBox.Font = 'Microsoft Sans Serif,10'

$DeptLabel = New-Object system.Windows.Forms.Label
$DeptLabel.text = "Department"
$DeptLabel.AutoSize = $false
$DeptLabel.width = 77
$DeptLabel.height = 20
$DeptLabel.location = New-Object System.Drawing.Point(12, 119)
$DeptLabel.Font = 'Microsoft Sans Serif,10'

$DeptComboBox = New-Object system.Windows.Forms.ComboBox
$DeptComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$DeptComboBox.Sorted = $true
$DeptComboBox.width = 110
$DeptComboBox.height = 20
@('Accounting', 'Administration', 'Assembly', 'Beauty', 'Dental', 'Factory', 'HR', 'IT', 'Marketing', 'Production', 'Quality', 'Sales/Design') | ForEach-Object {[void] $DeptComboBox.Items.Add($_)}
$DeptComboBox.location = New-Object System.Drawing.Point(90, 117)
$DeptComboBox.Font = 'Microsoft Sans Serif,10'

$LocationLabel = New-Object system.Windows.Forms.Label
$LocationLabel.text = "Location"
$LocationLabel.AutoSize = $false
$LocationLabel.width = 57
$LocationLabel.height = 20
$LocationLabel.location = New-Object System.Drawing.Point(210, 119)
$LocationLabel.Font = 'Microsoft Sans Serif,10'

$LocationComboBox = New-Object system.Windows.Forms.ComboBox
$LocationComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$LocationComboBox.Sorted = $true
$LocationComboBox.width = 110
$LocationComboBox.height = 20
@('California', 'Dental', 'Illinois', 'Kansas', 'Massachusetts', 'Maryland', 'Missouri', 'North Carolina', 'New Jersey', 'New York', 'Texas') | ForEach-Object {[void] $LocationComboBox.Items.Add($_)}
$LocationComboBox.location = New-Object System.Drawing.Point(267, 117)
$LocationComboBox.Font = 'Microsoft Sans Serif,10'

$TitleLabel = New-Object system.Windows.Forms.Label
$TitleLabel.text = "Title"
$TitleLabel.AutoSize = $false
$TitleLabel.width = 32
$TitleLabel.height = 20
$TitleLabel.location = New-Object System.Drawing.Point(105, 144)
$TitleLabel.Font = 'Microsoft Sans Serif,10'

$TitleTextBox = New-Object system.Windows.Forms.TextBox
$TitleTextBox.multiline = $false
$TitleTextBox.width = 160
$TitleTextBox.height = 20
$TitleTextBox.location = New-Object System.Drawing.Point(138, 143)
$TitleTextBox.Font = 'Microsoft Sans Serif,8'

$CopyAccessCheckBox = New-Object system.Windows.Forms.CheckBox
$CopyAccessCheckBox.text = "Copy access from another user (optional)"
$CopyAccessCheckBox.AutoSize = $false
$CopyAccessCheckBox.width = 310
$CopyAccessCheckBox.height = 20
$CopyAccessCheckBox.location = New-Object System.Drawing.Point(75, 200)
$CopyAccessCheckBox.Font = 'Microsoft Sans Serif,8'

$CopyUserLabel = New-Object system.Windows.Forms.Label
$CopyUserLabel.text = "Name"
$CopyUserLabel.AutoSize = $false
$CopyUserLabel.width = 42
$CopyUserLabel.height = 20
$CopyUserLabel.location = New-Object System.Drawing.Point(90, 222)
$CopyUserLabel.Font = 'Microsoft Sans Serif,10'

$CopyUserTextBox = New-Object system.Windows.Forms.TextBox
$CopyUserTextBox.Enabled = $false
$CopyUserTextBox.multiline = $false
$CopyUserTextBox.width = 140
$CopyUserTextBox.height = 20
$CopyUserTextBox.location = New-Object System.Drawing.Point(133, 220)
$CopyUserTextBox.Font = 'Microsoft Sans Serif,10'

$LogTextBox = New-Object system.Windows.Forms.TextBox
$LogTextBox.Text = "Logs will be displayed here.."
$LogTextBox.multiline = $true
$LogTextBox.width = 370
$LogTextBox.height = 246
$LogTextBox.location = New-Object System.Drawing.Point(420, 31)
$LogTextBox.Font = 'Microsoft Sans Serif,8'

$SubmitButton = New-Object system.Windows.Forms.Button
$SubmitButton.text = "Submit"
$SubmitButton.width = 125
$SubmitButton.height = 30
$SubmitButton.location = New-Object System.Drawing.Point(75, 250)
$SubmitButton.Font = 'Microsoft Sans Serif,10,style=Bold'

$ClearButton = New-Object system.Windows.Forms.Button
$ClearButton.text = "Clear"
$ClearButton.width = 125
$ClearButton.height = 30
$ClearButton.location = New-Object System.Drawing.Point(220, 250)
$ClearButton.Font = 'Microsoft Sans Serif,10,style=Bold'

$NewEmployeeForm.controls.AddRange(@($ITOnboardLabel, $UserInfoGroup, $CopyAccessCheckBox, $CopyUserLabel, $CopyUserTextBox, $SubmitButton, $ClearButton, $LogTextBox))
$UserInfoGroup.controls.AddRange(@($FirstNameLabel, $FirstNameTextBox, $LastNameLabel, $LastNameTextBox, $EmailLabel, $EmailTextBox, $EmailSuffixLabel, $PasswordLabel, $PasswordMaskedTextBox, $DeptLabel, $DeptComboBox, $LocationLabel, $LocationComboBox, $TitleLabel, $TitleTextBox))

#endregion GUI }

#region gui events {

$LastNameTextBox.Add_TextChanged( {
        $EmailTextBox.Text = ($FirstNameTextBox.Text.Substring(0, 1) + $LastNameTextBox.Text).ToLower();
    })

$CopyAccessCheckBox.Add_CheckedChanged( {  
        $CopyUserTextBox.Enabled = $CopyAccessCheckBox.Checked;
    })

$ClearButton.Add_Click( {
        clearFields
    })

$SubmitButton.Add_Click( {
        #Clear logs
        if ($LogTextBox.Text) {
            $LogTextBox.Text = ""
        }

        #Assign variables for user data
        $FirstName = $FirstNameTextBox.Text
        $LastName = $LastNameTextBox.Text
        $FullName = $FirstName + " " + $LastName
        $UserName = $EmailTextBox.Text

        # Get password
        $Password = $PasswordMaskedTextBox.Text | ConvertTo-SecureString -AsPlainText -Force

        # Create Email name
        $Domain = "@takarabelmont.com"
        $Email = $UserName + $Domain
    
        $Department = $DeptComboBox.SelectedItem
        $Location = $LocationComboBox.SelectedItem
        $Title = $TitleTextBox.Text

        if ($CopyAccessCheckBox.Checked -eq $true) {
            $CopyUser = $CopyUserTextBox.Text
        }
    
        # Department and Location fields should be selected before continuing
        if ($Department -eq $null -or $Location -eq $null) {
            [System.Windows.Forms.MessageBox]::Show('Department/Location is not selected', 'Invalid selection', 'OK', 'Error')
        }
        # Determine OU to place new user
        else {
            If ($Location -eq "New Jersey") {
                Switch ( $Department ) {
                    Accounting {$OU = "OU=Accounting,OU=NJ Users,OU=NJ - HQ,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"}
                    Administration {$OU = "OU=Administration,OU=NJ Users,OU=NJ - HQ,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"}
                    Assembly {$OU = "OU=Assembly,OU=NJ Users,OU=NJ - HQ,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"}
                    Beauty {$OU = "OU=Beauty,OU=NJ Users,OU=NJ - HQ,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"}
                    Dental {$OU = "OU=Dental,OU=NJ Users,OU=NJ - HQ,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"}
                    Factory {$OU = "OU=Factory,OU=NJ Users,OU=NJ - HQ,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"}
                    HR {$OU = "OU=HR,OU=NJ Users,OU=NJ - HQ,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"}
                    IT {$OU = "OU=IT,OU=NJ Users,OU=NJ - HQ,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"}
                    Marketing {$OU = "OU=Marketing,OU=NJ Users,OU=NJ - HQ,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"}
                    Production {$OU = "OU=Production,OU=NJ Users,OU=NJ - HQ,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"}
                    Quality {$OU = "OU=Quality,OU=NJ Users,OU=NJ - HQ,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"}
                    Sales/Design {$OU = "OU=NJ Users,OU=NJ,OU=Showrooms,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"} 
                }
                $HomeFolder = "\\takarabelmont.com\USER\NJ-HOME\$Username"
            }
            elseif ($Location -eq "Missouri" -and $Department -ne "Sales/Design") {
                $OU = "OU=Koken Users,OU=Koken,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"
                $HomeFolder = "\\takarabelmont.com\USER\MO-HOME\$Username"
            }
            else {
                Switch ( $Location ) {
                    'California' {
                        $OU = "OU=CA Users,OU=CA,OU=Showrooms,OU=TB,OU=Business Units,DC=takarabelmont,DC=com" 
                        $HomeFolder = "\\takarabelmont.com\USER\CA-HOME\$Username"
                    }
                    'Dental' {
                        $OU = "OU=Dental Users,OU=Dental,OU=Showrooms,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"
                        $HomeFolder = "\\takarabelmont.com\USER\NJ-HOME\$Username"
                    }
                    'Illinois' {
                        $OU = "OU=IL Users,OU=IL,OU=Showrooms,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"
                        $HomeFolder = "\\takarabelmont.com\USER\IL-HOME\$Username"
                    }
                    'Kansas' {
                        $OU = "OU=Marble Users,OU=Marble,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"
                        $HomeFolder = "\\takarabelmont.com\USER\KS-HOME\$Username"
                    }
                    'Massachusetts' {
                        $OU = "OU=MA Users,OU=MA,OU=Showrooms,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"
                        $HomeFolder = "\\takarabelmont.com\USER\NJ-HOME\$Username"
                    }
                    'Maryland' {
                        $OU = "OU=MD Users,OU=MD,OU=Showrooms,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"
                        $HomeFolder = "\\takarabelmont.com\USER\NJ-HOME\$Username"
                    }
                    'Missouri' {
                        $OU = "OU=MO Users,OU=MO,OU=Showrooms,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"
                        $HomeFolder = "\\takarabelmont.com\USER\MO-HOME\$Username"
                    }
                    'North Carolina' {
                        $OU = "OU=NC Users,OU=NC,OU=Showrooms,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"
                        $HomeFolder = "\\takarabelmont.com\USER\NJ-HOME\$Username"
                    }
                    'New York' {
                        $OU = "OU=NY Users,OU=NY,OU=Showrooms,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"
                        $HomeFolder = "\\takarabelmont.com\USER\NY-HOME\$Username"
                    }
                    'Texas' {
                        $OU = "OU=TX Users,OU=TX,OU=Showrooms,OU=TB,OU=Business Units,DC=takarabelmont,DC=com"
                        $HomeFolder = "\\takarabelmont.com\USER\TX-HOME\$Username"
                    }
                }
            }

            # Create user account
            try {
                New-ADUser -SamAccountName $UserName -AccountPassword $Password -CannotChangePassword $false -ChangePasswordAtLogon $true `
                    -Department $Department -EmailAddress $Email -Enabled $true -HomeDrive "Z:" -HomeDirectory $HomeFolder `
                    -Name $FullName -GivenName $FirstName -Surname $LastName -DisplayName $FullName `
                    -UserPrincipalName $UserName@takarabelmont.com -Title $Title -Description $Title -OfficePhone "TBD"

                $LogTextBox.ForeColor = 'black'
                $LogTextBox.Text += "Account for $Fullname was created successfully." + "`r`n"

                # Copy group membership of a user
                if ($CopyUser) {
                    $CopyGroups = Get-ADUser -Filter {(Name -eq $CopyUser) -or (SamAccountName -eq $CopyUser)} -SearchBase 'OU=TB,OU=Business Units,DC=takarabelmont,DC=com' -Properties memberof | select -Expand memberof
                    if($CopyGroups){
                        $CopyGroups | Add-ADGroupMember -Members $UserName
                        $LogTextBox.Text += "Access group membership copied from $CopyUser." + "`r`n"
                    }
                    else{
                        $LogTextBox.Text += "Failed to copy access from $CopyUser - User or Groups not found." + "`r`n"
                    }
                }

                #Move new user to the correct OU
                $User = Get-ADUser $UserName
                if ($User) {
                    $User | Move-ADObject -TargetPath $OU
                    $LogTextBox.Text += "$Fullname's account was moved to the $Location - $Department OU." + "`r`n"

                    $EmailMsg = "<font color=`"black`"><h4>New User Created - $Fullname</h4></font>"
                    $EmailMsg += "Please forward this information to the appropriate department.<br>"
                    $EmailMsg += "IT will need to ensure a GP account is created and O365 permissions are set.<br>"                
                       
                    Send-Email -emailSubject "New User - $Fullname" -htmlBody $EmailMsg
                    $LogTextBox.Text += "Email confirmation sent to HR and IT" + "`r`n"
                    Invoke-Command NJ-vSYNC-01 -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
                }
            }
            catch {
                $LogTextBox.ForeColor = 'red'
                $ErrorMessage = $_.Exception.Message
                $LogTextBox.Text = $ErrorMessage

                $EmailMsg = "<font color=`"black`"><h4>Creation Error - $Fullname</h4></font>"
                $EmailMsg += "There was a error creating an account for $Fullname, see logs and report below.<br>"
                Send-Email -emailSubject "ERROR Creating User - $Fullname" -htmlBody $EmailMsg
            }
        }
    })
#endregion events }

function clearFields {
    $FirstNameTextBox.Text = ""
    $LastNameTextBox.Text = ""
    $EmailTextBox.Text = ""
    $PasswordMaskedTextBox.Text = ""
    $DeptComboBox.SelectedItem = $null
    $LocationComboBox.SelectedItem = $null
    $TitleTextBox.Text = ""
    $LogTextBox.Text = ""
}

function Send-Email($emailSubject, $htmlBody) {
    # Table name
    $tableName = "New User Report"

    #Create Table object
    $table = New-Object system.Data.DataTable "$tableName"

    #Define Columns
    $col = New-Object system.Data.DataColumn "$Fullname",([string])

    #Add the Columns
    $table.columns.add($col)

    #Create rows
    $row1 = $table.NewRow()
    $row1.$Fullname = $Title
    $table.Rows.Add($row1)

    $row2 = $table.NewRow()
    $row2.$Fullname = $Email
    $table.Rows.Add($row2)

    $row3 = $table.NewRow()
    $row3.$Fullname = "Username: $UserName"
    $table.Rows.Add($row3)

    #Get password from secure string
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

    #Add password to table
    $row4 = $table.NewRow()
    $row4.$Fullname = "Password: $UnsecurePassword"
    $table.Rows.Add($row4)

    #Clear variable containing password
    Clear-Variable -Name UnsecurePassword

#Creating head and table style
$Head = @"
      
<style>
body {
    font-family: "Arial";
    font-size: 9pt;
    color: #000000;
    }
th, td { 
    border: 1px solid #e57300;
    border-collapse: collapse;
    padding: 5px;
    }
th {
    font-size: 1.2em;
    text-align: left;
    background-color: #003366;
    color: #ffffff;
    }
td {
    color: #000000;
    }
.even { background-color: #ffffff; }
.odd { background-color: #bfbfbf; }
</style>
      
"@

    $body = $htmlBody
    $body += "<font color=`"red`">** Replies to this email will create a helpdesk ticket **</font><br><br>"

    # Creating body
    [string]$body = [PSCustomObject]$table | ConvertTo-HTML -Property $Fullname -head $head -Body $body

    #Add logs to the end of body
    $body += "<br><br><br><font color=`"#4C607B`">"
    $body += "--- IT Logs ---<br>"
    $logs = $LogTextBox.Text.Split('.').Trim()
    foreach($log in $logs){
        $body += $log + "<br>"
    }

    # Setup email parameters
    $subject = $emailSubject
    $priority = "High"
    $smtpServer = "NJ-vSMTP-01"
    $emailFrom = "helpdesk@takarabelmont.com"
    $emailTo = "ppatel@takarabelmont.com"

    # Send the report email
    Send-MailMessage -To $emailTo -Subject $subject -BodyAsHtml $body -SmtpServer $smtpServer -From $emailFrom -Priority $priority
}

#Run code
[void]$NewEmployeeForm.ShowDialog()