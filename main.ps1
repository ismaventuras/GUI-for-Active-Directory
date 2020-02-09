<#Global variables#>
$ou = "OU=ES,DC=EU,DC=DIR,DC=BUNGE,DC=COM"


function RegisterAction {
    param($string)
    $date=get-date -Format '[hh:mm dd:mm:yy]'
    $Textbox_Register.AppendText([System.Environment]::Newline)
    $registerString = '{0} {1}' -f $date , $string
    $Textbox_Register.AppendText($registerString)
}

function UnlockUser {
    param ( $userName
    )
    if($userListComboBox.selectedindex -ne -1){
    Enable-ADAccount $userListComboBox.selecteditem
    $outString = 'Account has been unlocked for {0}' -f $userListComboBox.selecteditem
    RegisterAction $outString
    }
    else {
        [System.Windows.MessageBox]::Show('No user selected.')
    }
}
function GetADGroups {
    if($userListComboBox.selectedindex -ne -1){
        $title='AD Groups for {0}' -f $userListComboBox.SelectedValue
        RegisterAction  $title
        foreach($group in (Get-ADPrincipalGroupMembership $userListComboBox.selecteditem).name){
            RegisterAction $group
        }
        $Textbox_Register.AppendText([System.Environment]::Newline)
    }
    else {
        [System.Windows.MessageBox]::Show('No user selected.')
    }
}
function ResetPasswordAD {
    if($userListComboBox.selectedindex -ne -1){
    Set-ADAccountPassword -Identity $userListComboBox.selecteditem -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "Bunge2020" -Force) -PassThru -Confirm:$false
    Set-ADUser -ChangePasswordAtLogon $true -Identity $userListComboBox.selecteditem -Confirm:$false
    UnlockUser
    $outString = 'Password has been reset to default for user {0}' -f $userListComboBox.selecteditem
    RegisterAction $outString
    }
    else {
        [System.Windows.MessageBox]::Show('No user selected.')
    }
    }

function FillUserList {
    $userListComboBox.ItemsSource = get-aduser -f * | sort -property Name
    $userListComboBox.DisplayMemberPath="Name"
}
function OpenExplorer{
    if($computerListComboBox.selectedindex -ne -1){
        $path = "\\{0}" -f $computerListComboBox.SelectedValue
        write-host $computerListComboBox.selecteditem
        write-host $path
        #[System.Windows.MessageBox]::Show()
        explorer $path
    }
    else {
        [System.Windows.MessageBox]::Show('No user selected.')
    }
}
function FillComputerList{
    $computerListComboBox.ItemsSource = get-adcomputer -f * -properties description  |sort -property Description| select description,name 
    $computerListComboBox.DisplayMemberPath="description"
    $computerListComboBox.SelectedValuePath="name"

}
#Library to make the GUI work
Add-Type -AssemblyName PresentationFramework
#GUI code
[xml]$XAML = @'
<Window x:Name="Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AD User tool" Height="343.664" Width="801.541">
    <TabControl x:Name="TabControlMain" HorizontalAlignment="Left" Height="311" VerticalAlignment="Top" Width="792">
        <TabItem x:Name="headerUsers" Header="Users">
            <Grid x:Name="mainGrid" Width="794" Margin="0,0,-8,0">
                <Label x:Name="userListBoxLabel" Content="Select a user" HorizontalAlignment="Left" Margin="11,19,0,0" VerticalAlignment="Top"/>
                <ComboBox x:Name="userListComboBox" HorizontalAlignment="Left" Margin="92,23,0,0" VerticalAlignment="Top" Width="120" IsEditable="True">

                </ComboBox>
                <Button x:Name="button_UnlockUser" Content="Unlock user" HorizontalAlignment="Left" Margin="12,75,0,0" VerticalAlignment="Top" Width="91" Height="61"/>
                <Button x:Name="button_ResetPassword" Content="Reset Password" HorizontalAlignment="Left" Margin="108,75,0,0" VerticalAlignment="Top" Width="91" Height="61"/>
                <Button x:Name="button_GetGroups" Content="Get groups" HorizontalAlignment="Left" Margin="204,75,0,0" VerticalAlignment="Top" Width="91" Height="61"/>
                <TextBox x:Name="Textbox_Register" HorizontalAlignment="Left" Height="83" Margin="0,200,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="794"/>

            </Grid>
        </TabItem>
        <TabItem x:Name="headerComputers" Header="Computers">
            <Grid x:Name="Grid_computers" Width="794" Margin="0,0,-8,0">
                <Label x:Name="computerListBoxLabel" Content="Select a computer" HorizontalAlignment="Left" Margin="0,5,0,0" VerticalAlignment="Top"/>
                <ComboBox x:Name="computerListComboBox" HorizontalAlignment="Left" Margin="10,31,0,0" VerticalAlignment="Top" Width="174" IsEditable="True">

                </ComboBox>
                <TextBox x:Name="Textbox_RegisterComputer" HorizontalAlignment="Left" Height="83" Margin="0,200,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="794"/>
                <Button x:Name="button_OpenExplorer" Content="Open Explorer" HorizontalAlignment="Left" Margin="10,58,0,0" VerticalAlignment="Top" Width="95" Height="42"/>

            </Grid>
        </TabItem>
    </TabControl>
</Window>
'@
#Create a XML reader and read XAML
$reader=(New-Object System.Xml.XmlNodeReader $XAML)
#Load  XAML into a variable using XAML reader (without the assembly on the first line the program won't load)
$GUI=[Windows.Markup.XAMLreader]::Load($reader)
#Reading controls in the WPF XAML and creating variables with the names of the controls to work with them.
$XAML.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object {New-Variable  -Name $_.Name -Value $GUI.FindName($_.Name) -Force}

#Filling user and computer list 
FillUserList
FillComputerList

#Handling Buttons
$button_UnlockUser.Add_Click({ UnlockUser})
$button_GetGroups.Add_Click({ GetADGroups})
$button_ResetPassword.Add_Click({ResetPasswordAD})
$button_OpenExplorer.Add_Click({OpenExplorer})



#Show the dialog
$Null = $GUI.ShowDialog() 
##END OF GUI SCRIPT

