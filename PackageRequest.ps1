#Xaml to create the form
[xml]$form = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="PDQ Package Request" Height="450" Width="704.929" Background="#FFF0F0F0">
    <Grid Name="ManagerLabel" Margin="2,-6,-157,4">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
        </Grid.ColumnDefinitions>
        <TextBox Name="AppNameTextbox" HorizontalAlignment="Left" Height="23" Margin="446,72,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="212" TabIndex="4"/>
        <TextBox Name="AppVersionTextbox" HorizontalAlignment="Left" Height="23" Margin="103,111,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="116" TabIndex="5"/>
        <TextBox Name="LocationTextbox" HorizontalAlignment="Left" Height="23" Margin="19,263,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="285" TabIndex="8"/>
        <TextBox Name="FirstNameTextbox" HorizontalAlignment="Left" Height="23" Margin="103,29,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="212" TabIndex="1"/>
        <Label Name="AppLocationLabel" Content="If we install a previous version where is the installer located?" HorizontalAlignment="Left" Margin="19,231,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
        <Label Name="FirstNameLabel" Content="First Name" HorizontalAlignment="Left" Margin="10,29,0,0" VerticalAlignment="Top" FontWeight="Bold" Height="23" Width="77"/>
        <Button Name="SubmitButton" Content="Submit" HorizontalAlignment="Left" Margin="19,374,0,0" VerticalAlignment="Top" Width="285" Height="33" TabIndex="11" ClickMode="Press"/>
        <TextBox Name="CommentsTextBox" HorizontalAlignment="Left" Height="52" Margin="19,312,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="285" AcceptsTab="True" TabIndex="9"/>
        <Label Name="CommentsLabel" Content="Additional comments (if any)" HorizontalAlignment="Left" Margin="19,286,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
        <Label Name="AppNameLabel" Content="Application Name" HorizontalAlignment="Left" Margin="329,72,0,0" VerticalAlignment="Top" Height="29" FontWeight="Bold" Width="112"/>
        <Label Name="AppVersionLabel" Content="Version" HorizontalAlignment="Left" Margin="10,111,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.145,0.466" FontWeight="Bold"/>
        <RadioButton Name="UsersRadio" Content="Multiple Users" HorizontalAlignment="Left" Margin="19,169,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.137,-0.274" Width="113" Height="19" TabIndex="7"/>
        <RadioButton Name="DepartmentsRadio" Content="Multiple Departments" HorizontalAlignment="Left" Margin="19,193,0,0" VerticalAlignment="Top" Width="140" Height="18" TabIndex="7"/>
        <RadioButton Name="CompanyRadio" Content="Company Wide" HorizontalAlignment="Left" Margin="19,216,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.5,0.5" Width="109" Height="20" TabIndex="7">
            <RadioButton.RenderTransform>
                <TransformGroup>
                    <ScaleTransform/>
                    <SkewTransform AngleX="-4.161"/>
                    <RotateTransform/>
                    <TranslateTransform X="0.799"/>
                </TransformGroup>
            </RadioButton.RenderTransform>
        </RadioButton>
        <Label Name="AppScopeLabel" Content="Deployment Scope" HorizontalAlignment="Left" Margin="19,144,0,0" VerticalAlignment="Top" Height="31" Width="137" FontWeight="Bold"/>
        <Label Name="DateLabel" Content="Package needed by date" HorizontalAlignment="Left" Margin="475,227,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.145,0.466" Height="29" Width="153" FontWeight="Bold"/>
        <Calendar Name="Calendar" HorizontalAlignment="Left" Margin="446,252,0,0" VerticalAlignment="Top" Width="212" TabIndex="10" Height="165"/>
        <TextBox Name="YourEmailTextBox" HorizontalAlignment="Left" Height="23" Margin="103,70,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="212" TabIndex="3"/>
        <Label Name="YourEmailLabel" Content="Email address" HorizontalAlignment="Left" Margin="5,70,0,0" VerticalAlignment="Top" Height="31" FontWeight="Bold" Width="93" RenderTransformOrigin="0.529,1.052"/>
        <TextBox Name="LastNameTextbox" HorizontalAlignment="Left" Height="23" Margin="446,29,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="212" TabIndex="2"/>
        <Label Name="LastNameLabel" Content="Last Name" HorizontalAlignment="Left" Margin="329,29,0,0" VerticalAlignment="Top" FontWeight="Bold" Height="23" Width="77"/>
        <TextBox Name="JustificationTextbox" HorizontalAlignment="Left" Height="111" Margin="446,112,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="212" AcceptsTab="True" TabIndex="6"/>
        <Label Name="JustificationLabel" Content="Justification" HorizontalAlignment="Left" Margin="329,111,0,0" VerticalAlignment="Top" FontWeight="Bold" Height="27" Width="85"/>
        <ComboBox Name="ManagerComboBox" HorizontalAlignment="Left" Margin="195,175,0,0" VerticalAlignment="Top" Width="120">
            <ComboBoxItem Content="Kathy Baumgart"/>
            <ComboBoxItem Content="Jason Conner"/>
            <ComboBoxItem Content="Amos Mincey"/>
        </ComboBox>
        <Label Name="AppScopeLabel_Copy" Content="Manager" HorizontalAlignment="Left" Margin="195,144,0,0" VerticalAlignment="Top" Height="31" Width="120" FontWeight="Bold"/>

    </Grid>
</Window>
"@

#Add in the PresentationFramwork Class
Add-Type -AssemblyName PresentationFramework

#Create an XML node reader
$NR = (New-Object System.Xml.XmlNodeReader $Form)

#$Create a XAML Reader and load the XML in
$Win = [Windows.Markup.XamlReader]::Load($NR)

#Initialize a variable for each form control
$FirstName = $Win.FindName("FirstNameTextbox")
$LastName = $Win.FindName("LastNameTextbox")
$UserEmail = $Win.FindName("YourEmailTextBox")
$AppName = $Win.FindName("AppNameTextbox")
$AppVersion = $Win.FindName("AppVersionTextbox")
$Location = $Win.FindName("LocationTextbox")
$Comments = $Win.FindName("CommentsTextBox")
$Justification = $Win.FindName("JustificationTextbox")
$Date = $Win.FindName("Calendar")
$Users = $Win.FindName("UsersRadio")
$Departments = $Win.FindName("DepartmentsRadio")
$Company = $Win.FindName("CompanyRadio")
$Submit = $Win.FindName("SubmitButton")
$Manager = $Win.FindName("ManagerComboBox")

#Determine which radio button is checked
$Radios = @(
    $Users
    $Departments
    $Company
)

$Regex = '^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$'
#Create a function for determining if any of the manadatory fields were left black and return an associated error message

function Return-Error {
    $Script:Checked = $radios.GetEnumerator() | Where-Object ischecked -eq $True | Select-Object -ExpandProperty Content
    if ($FirstName.Text.Length -lt 1) {
        return "The First Name field is blank.  It is a required field."
        break
    }
    elseif ($LastName.Text.Length -lt 1) {
        return "The Last Name field is blank.  It is a required field."
        break
    }
    elseif ($UserEmail.Text.Length -lt 1) {
        return "The email field is blank.  It is a required field."
        break
    }
    elseif ($UserEmail.Text -notmatch $Regex) {
        return "The email address entered is invalid."
    }
    elseif ($AppName.Text.Length -lt 1) {
        return "The application name field is blank.  It is a required field."
        break
    }
    elseif ($AppVersion.Text.Length -lt 1) {
        return "The application version field is blank.  It is a required field."
        break
    }
    elseif ($Justification.Text.Length -lt 1) {
        return "The justification field is blank.  It is a required field."
        break
    }
    elseif (!($Date.SelectedDate)) {
        return "The date field is blank.  It is a required field."
        break
    }
    elseif ($Checked.Length -lt 1) {
        return "No scope was selected.  Please try again."
        break
    }
    elseif ($Manager.text.Length -lt 1) {
        return "A manager was not selected.  This is a required field"
        break
    }
    else {
        return "Passed"
    }
}

#Associate the selected manager to their email address to send them an email as well
$ManagerEmail = switch ($Manager.text)
{
    'Manager 1' {'Manager1@email.com'}
    'Manager 2' {'Manager1@email.com'}
    'Mr. Manager' {'MrManager@email.com'}
}

$emails = @(
"PDQ package builder manager email"
$ManagerEmail
$UserEmail
)

#Create a function for emailing the form to the user and the package approver
function Send-Form {
param ($To)
    $Params = @{
        SmtpServer = 'secretemailserver'
        To         = $To
        From       = 'secretEmail'
        Subject    = "PDQ Package Build Request for $($AppName.text)"
        Body       = @"
    A request has been generated for a PDQ package.  Please see below.

    Requested by: $($FirstName.text + ' ' + $Lastname.text)
    Application Name: $($AppName.text)
    Application Version: $($AppVersion.text)
    Date package is needed by: $($Date.SelectedDate.ToShortDateString())
    Justification: $($Justification.text)
    Scope: $Checked
"@
    }
    Send-MailMessage @Params
}

$Submit.add_click( {
        $Return = Return-Error
        if ($Return -ne "Passed") {
            [System.Windows.MessageBox]::Show($Return)
        }
        else {
            foreach ($email in $emails) {
                Send-Form -To $email
            }
            $Win.Close()
        }
    })

[void]$Win.ShowDialog()