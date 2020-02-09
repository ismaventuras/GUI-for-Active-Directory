
function RegisterAction {
    param($string)
    $date=get-date -Format '[hh:mm dd:mm:yy]'
    $Textbox_Register.AppendText([System.Environment]::Newline)
    $registerString = '{0} {1}' -f $date , $string
    $Textbox_Register.AppendText($registerString)
}

function UnlockUser {

<#
    .Synopsis
     Unlocks the account for the selected user
    .DESCRIPTION
    Uses Unlock-ADAccount to unlock the specified user account that has been blocked after several attempts trying to login.
#>
    param ( $userName
    )
    Enable-ADAccount $userListComboBox.selecteditem
    [System.Windows.MessageBox]::Show('Account has been unlocked '+ $userListComboBox.selecteditem);
}

function GetADGroups {
   # $Textbox_Register.Text += 'AD Groups', $userListComboBox.SelectedValue ,' is member of:1&#x0a;'
    #$Textbox_Register.Text += 'AD Groups for {0} `r`n' -f $userListComboBox.SelectedValue
    $title='AD Groups for {0}' -f $userListComboBox.SelectedValue
    RegisterAction  $title
    #$Textbox_Register.AppendText('AD Groups for {0}' -f $userListComboBox.SelectedValue)
    #$Textbox_Register.AppendText([System.Environment]::Newline)
    foreach($group in (Get-ADPrincipalGroupMembership $userListComboBox.selecteditem).name){
        RegisterAction $group
        #$Textbox_Register.AppendText([System.Environment]::Newline)
        #$Textbox_Register.AppendText($group)
    }
    #$Textbox_Register.AppendText((Get-ADPrincipalGroupMembership $userListComboBox.selecteditem).name)
    $Textbox_Register.AppendText([System.Environment]::Newline)
}


function FillList {
    $userListComboBox.ItemsSource = get-aduser -f *
    $userListComboBox.DisplayMemberPath="Name"
}
#Library to make the GUI work
Add-Type -AssemblyName PresentationFramework
#GUI code
[xml]$XAML = @'
<Window x:Name="Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AD User tool" Height="343.664" Width="801.541">
    <Grid x:Name="mainGrid">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
        </Grid.ColumnDefinitions>
        <Label x:Name="userListComboBoxLabel"  Content="Select a user" HorizontalAlignment="Left" Margin="11,19,0,0" VerticalAlignment="Top"/>
        <ComboBox x:Name="userListComboBox" HorizontalAlignment="Left" Margin="92,23,0,0" VerticalAlignment="Top" Width="120">
        </ComboBox>
        <Button x:Name="button_UnlockUser" Content="Unlock user" HorizontalAlignment="Left" Margin="246,23,0,0" VerticalAlignment="Top" Width="75"/>
        <Button x:Name="button_ResetPassword" Content="Reset Password&#xD;&#xA;" HorizontalAlignment="Left" Margin="326,23,0,0" VerticalAlignment="Top" Width="91" Height="22"/>
        <Button x:Name="button_GetGroups" Content="Get groups&#xD;&#xA;" HorizontalAlignment="Left" Margin="433,23,0,0" VerticalAlignment="Top" Width="74" Height="22"/>
        <TextBox x:Name="Textbox_Register" HorizontalAlignment="Left" Height="83" Margin="0,230,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="794"/>

    </Grid>
</Window>
'@
#Create a XML reader and read XAML
$reader=(New-Object System.Xml.XmlNodeReader $XAML)
#Load  XAML into a variable using XAML reader (without the assembly on the first line the program won't load)
$GUI=[Windows.Markup.XAMLreader]::Load($reader)
#Reading controls in the WPF XAML and creating variables with the names of the controls to work with them.
$XAML.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object {New-Variable  -Name $_.Name -Value $GUI.FindName($_.Name) -Force}

#Filling user list 
FillList
$Textbox_Register.IsReadOnly = $true;
$Textbox_Register.AcceptsReturn = $true;
#Handling Buttons
$button_UnlockUser.Add_Click({ UnlockUser})
$button_GetGroups.Add_Click({ GetADGroups})




#Show the dialog
$Null = $GUI.ShowDialog() 
##END OF GUI SCRIPT

