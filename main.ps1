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
    <#    $title='AD Groups for {0}' -f $userListComboBox.SelectedValue
        RegisterAction  $title
        foreach($group in (Get-ADPrincipalGroupMembership $userListComboBox.selecteditem).name){
            RegisterAction $group
        }
        $Textbox_Register.AppendText([System.Environment]::Newline)#>
        Get-ADPrincipalGroupMembership $userListComboBox.selecteditem | out-gridview
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
function StartSCCMRemote{
    if($computerListComboBox.selectedindex -ne -1){
        $path = ".\SCCM\CmRc.exe {0}" -f $computerListComboBox.SelectedValue
        write-host $computerListComboBox.selecteditem
        write-host $path
        #[System.Windows.MessageBox]::Show()
        explorer $path
    }
    else {
        [System.Windows.MessageBox]::Show('No user selected.')
    }
}
function isUp{

    $online=Test-Connection $computerListComboBox.SelectedValue -Quiet -Count 1

    if($online){
        $labelConnectivity.Content = 'online'
        $labelConnectivity.Background = '#32CD32'
    }
    else {
        $labelConnectivity.Content = 'offline'
        $labelConnectivity.Background = '#e50000'
    }
}
function FillComputerList{
    $computerListComboBox.ItemsSource = get-adcomputer -f * -properties description  |sort -property Description| select description,name 
    $computerListComboBox.DisplayMemberPath="description"
    $computerListComboBox.SelectedValuePath="name"

}

function creds{
    $credPath = ".\creds.xml"
    # check for stored credential
    if ( Test-Path $credPath ) {
        #crendetial is stored, load it 
        $cred = Import-CliXml -Path $credPath
        $cred
    } else {
        # no stored credential: create store, get credential and save it
        $parent = split-path $credpath -parent
        if ( -not (test-Path $parent)) {
            New-Item -ItemType Directory -Force -Path $parent
        }
        $cred = get-credential -Message "Please use an account with rights on AD"
        $cred | Export-CliXml -Path $credPath
        $cred
    }
}

#Library to make the GUI work
Add-Type -AssemblyName PresentationFramework
#GUI code
[xml]$XAML = @'
<Window x:Name="Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AD User tool" Height="343.664" Width="801.541">
        <Grid HorizontalAlignment="Left" Height="344" VerticalAlignment="Top" Width="802">
        <TabControl x:Name="TabControlMain" HorizontalAlignment="Left" Height="344" VerticalAlignment="Top" Width="802">
            <TabItem x:Name="headerUsers" Header="Users">
                <Grid x:Name="mainGrid" Width="794" Margin="0,0,-8,0">
                    <Label x:Name="userListBoxLabel" Content="Select a user" HorizontalAlignment="Left" Margin="10,5,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="userListComboBox" HorizontalAlignment="Left" Margin="10,31,0,0" VerticalAlignment="Top" Width="185" IsEditable="True">

                    </ComboBox>
                    <Button x:Name="button_UnlockUser" Content="Unlock user" HorizontalAlignment="Left" Margin="10,58,0,0" VerticalAlignment="Top" Width="91" Height="61"/>
                    <Button x:Name="button_ResetPassword" Content="Reset Password" HorizontalAlignment="Left" Margin="10,119,0,0" VerticalAlignment="Top" Width="91" Height="61"/>
                    <Button x:Name="button_GetGroups" Content="Get groups" HorizontalAlignment="Left" Margin="106,58,0,0" VerticalAlignment="Top" Width="91" Height="61"/>
                    <TextBox x:Name="Textbox_Register" HorizontalAlignment="Left" Height="83" Margin="0,202,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="784"/>

                </Grid>
            </TabItem>
            <TabItem x:Name="headerComputers" Header="Computers">
                <Grid x:Name="Grid_computers" Width="794" Margin="0,0,-8,0">
                    <Label x:Name="computerListBoxLabel" Content="Select a computer" HorizontalAlignment="Left" Margin="10,5,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="computerListComboBox" HorizontalAlignment="Left" Margin="10,31,0,0" VerticalAlignment="Top" Width="366" IsEditable="True">

                    </ComboBox>
                    <TextBox x:Name="Textbox_RegisterComputer" HorizontalAlignment="Left" Height="83" Margin="0,191,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="768"/>
                    <Button x:Name="button_OpenExplorer" Content="Open Explorer" HorizontalAlignment="Left" Margin="10,88,0,0" VerticalAlignment="Top" Width="95" Height="42"/>
                    <Label x:Name="labelConnectivity" Content="" HorizontalAlignment="Left" Margin="381,31,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="button_Ping" Content="Ping" HorizontalAlignment="Left" Margin="110,88,0,0" VerticalAlignment="Top" Width="95" Height="42"/>
                    <Label x:Name="labelSelected" HorizontalAlignment="Left" Margin="10,53,0,0" VerticalAlignment="Top" Height="30" Width="142" FontSize="10"/>

                </Grid>
            </TabItem>
            <TabItem x:Name="headerOptions" Header="Options" HorizontalAlignment="Left" Height="20" VerticalAlignment="Top" Width="54">
                <Grid Background="#FFE5E5E5">
                    <Button x:Name="button_UpdateLists" Content="Update Lists" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="101" Height="43"/>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
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
$button_Ping.Add_Click({isUp})
$button_UpdateLists.Add_Click({FillUserList    
    FillComputerList})
#Ping on computer change
#$computerListComboBox.Add_SelectionChanged({isUp})
#Clean connectiviy status
$computerListComboBox.Add_SelectionChanged({$labelConnectivity.Background = $Null 
     $labelConnectivity.Content = ''
    $labelSelected.Content=$computerListComboBox.SelectedValue})


#Show the dialog
$Null = $GUI.ShowDialog() 
##END OF GUI SCRIPT

