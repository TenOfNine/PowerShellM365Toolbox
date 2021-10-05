<#
.SYNOPSIS
	Handy tool for stuff that usually requires PS, but I’m to lazy to use
.DESCRIPTION
    	Handy tool for stuff that usually requires PS, but I’m to lazy to use
.INPUTS
    	GUI
.OUTPUTS
  	None
.NOTES
	Version:        0.1.1
	Author:         TenOfNine
	Creation Date:  01.10.2021
.EXAMPLE
  	None provided
#>




using module Microsoft.PowerShell.Management
using module Microsoft.PowerShell.Utility

using namespace System.Management.Automation
using namespace System.PresentationFramework
using namespace System.Windows.Markup

$ErrorActionPreferenceBackup = $ErrorActionPreference
Set-StrictMode -Version 'Latest'

Add-Type -AssemblyName 'PresentationFramework'





$Script:My = [HashTable]::Synchronized(@{})
$Script:My.WindowXaml = @'
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PowerShellM365Toolbox"
        mc:Ignorable="d"
        Title="PowerShell M365 Toolbox" Height="490" Width="820" ResizeMode="NoResize" SizeToContent="Manual">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="17*"/>
            <ColumnDefinition Width="63*"/>
        </Grid.ColumnDefinitions>
        <TabControl Grid.ColumnSpan="2">
            <TabItem Header="Overview" Height="20">
                <Grid Background="#FFE5E5E5">
                    <GroupBox HorizontalAlignment="Left" Height="200" Header="To what am I connected?" Margin="10,10,0,0" VerticalAlignment="Top" Width="280">
                        <StackPanel>
                            <CheckBox x:Name="OverviewCBAzure" Content="Azure AD" IsEnabled="False"/>
                            <CheckBox x:Name="OverviewCBEXO" Content="Exchange Online" IsEnabled="False"/>
                            <CheckBox x:Name="OverviewCBTeams" Content="Teams" IsEnabled="False"/>
                        </StackPanel>
                    </GroupBox>
                    <GroupBox HorizontalAlignment="Left" Height="200" Header="About" Margin="587,10,0,0" VerticalAlignment="Top" Width="197">
                        <TextBlock HorizontalAlignment="Left" TextWrapping="Wrap" Text="PowerShell M365 Toolbox v0.1.1 01.10.2021&#10;&#13;https://github.com/TenOfNine/PowerShellM365Toolbox" Width="187"/>
                    </GroupBox>
                    <GroupBox HorizontalAlignment="Left" Height="194" Header="Required Modules" Margin="307,19,0,0" VerticalAlignment="Top" Width="260">
                        <TextBlock HorizontalAlignment="Left" TextWrapping="Wrap" Text="Install-Module Azure-AD" Width="250"/>
                    </GroupBox>
                </Grid>
            </TabItem>
            <TabItem Header="Azure AD" Height="20">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="27*"/>
                        <ColumnDefinition Width="767*"/>
                    </Grid.ColumnDefinitions>
                </Grid>
            </TabItem>
            <TabItem Header="EXO" Height="20">
                <Grid Background="#FFE5E5E5">
                    <TabControl>
                        <TabItem Header="Start">
                            <Grid Background="#FFE5E5E5">
                                <Button x:Name="exostartbtnconnectexo" Content="Connect to Exchange Online" HorizontalAlignment="Left" Height="30" Margin="14,19,0,0" VerticalAlignment="Top" Width="175"/>
                            </Grid>
                        </TabItem>
                        <TabItem Header="Roomlists">
                            <Grid Background="#FFE5E5E5">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="9*"/>
                                    <ColumnDefinition Width="799*"/>
                                </Grid.ColumnDefinitions>
                                <Button x:Name="exoroomlistsaddbtn" Content="Add" HorizontalAlignment="Left" Height="25" Margin="485,174,0,0" VerticalAlignment="Top" Width="40" Grid.Column="1" IsEnabled="False"/>
                                <Button x:Name="exoroomlistsdelbtn" Content="Del" HorizontalAlignment="Left" Height="25" Margin="485,208,0,0" VerticalAlignment="Top" Width="40" Grid.Column="1" IsEnabled="False"/>
                                <TextBox x:Name="exoroomlistsSearchtb1" HorizontalAlignment="Left" Height="20" Margin="0,39,0,0" TextWrapping="Wrap" Text="Search" VerticalAlignment="Top" Width="135" Grid.Column="1" IsEnabled="False"/>
                                <TextBox x:Name="exoroomlistsSearchtb2" HorizontalAlignment="Left" Height="20" Margin="539,39,0,0" TextWrapping="Wrap" Text="Search" VerticalAlignment="Top" Width="135" Grid.Column="1" IsEnabled="False"/>
                                <Button x:Name="exoroomlistsSearchbtn1" Content="Search" HorizontalAlignment="Left" Height="20" Margin="145,39,0,0" VerticalAlignment="Top" Width="86" Grid.Column="1" IsEnabled="False"/>
                                <Button x:Name="exoroomlistsSearchbtn2" Content="Search" HorizontalAlignment="Left" Height="20" Margin="683,39,0,0" VerticalAlignment="Top" Width="86" Grid.Column="1" IsEnabled="False"/>
                                <Label Content="Roomlists" HorizontalAlignment="Left" Height="24" Margin="61,10,0,0" VerticalAlignment="Top" Width="110" Grid.Column="1"/>
                                <Label Content="Rooms in selected roomlist" HorizontalAlignment="Left" Height="24" Margin="280,10,0,0" VerticalAlignment="Top" Width="160" Grid.Column="1"/>
                                <Label Content="All rooms" HorizontalAlignment="Left" Height="24" Margin="599,10,0,0" VerticalAlignment="Top" Width="110" Grid.Column="1"/>
                                <ListBox x:Name="exoroomlistslb1" Grid.Column="1" HorizontalAlignment="Left" Height="295" Margin="1,84,0,0" VerticalAlignment="Top" Width="230" d:ItemsSource="{d:SampleData ItemCount=5}"/>
                                <ListBox x:Name="exoroomlistslb2" Grid.Column="1" HorizontalAlignment="Left" Height="295" Margin="245,84,0,0" VerticalAlignment="Top" Width="230" d:ItemsSource="{d:SampleData ItemCount=5}"/>
                                <ListBox x:Name="exoroomlistslb3" Grid.Column="1" HorizontalAlignment="Left" Height="295" Margin="539,84,0,0" VerticalAlignment="Top" Width="230" d:ItemsSource="{d:SampleData ItemCount=5}"/>
                            </Grid>
                        </TabItem>
                    </TabControl>
                </Grid>
            </TabItem>
        </TabControl>

    </Grid>
</Window>




'@



$Script:My.Window = [XamlReader]::Parse($Script:My.WindowXaml)


$Script:My.CheckBox1 = $Script:My.Window.FindName('OverviewCBAzure')
$Script:My.CheckBox2 = $Script:My.Window.FindName('OverviewCBEXO')
$Script:My.CheckBox3 = $Script:My.Window.FindName('OverviewCBTeams')

$Script:My.ConnectExchangebtn = $Script:My.Window.FindName('exostartbtnconnectexo')

$Script:My.EXORoomlistView1 = $Script:My.Window.FindName('exoroomlistslb1')
$Script:My.EXORoomlistView2 = $Script:My.Window.FindName('exoroomlistslb2')
$Script:My.EXORoomlistView3 = $Script:My.Window.FindName('exoroomlistslb3')

$Script:My.EXORoomlistsAddBtn = $Script:My.Window.FindName('exoroomlistsaddbtn')
$Script:My.EXORoomlistsDelBtn = $Script:My.Window.FindName('exoroomlistsdelbtn')

$Script:My.EXORoomlistsSearchBtn1 = $Script:My.Window.FindName('exoroomlistsSearchbtn1')
$Script:My.EXORoomlistsSearchBtn2 = $Script:My.Window.FindName('exoroomlistsSearchbtn2')

$Script:My.EXORoomlistsSearchTb1 = $Script:My.Window.FindName('exoroomlistsSearchtb1')
$Script:My.EXORoomlistsSearchTb2 = $Script:My.Window.FindName('exoroomlistsSearchtb2')





$arraylistgroups = New-Object System.Collections.ArrayList 
$arraylistrooms = New-Object System.Collections.ArrayList 
$arraylistmembers = New-Object System.Collections.ArrayList 

$Script:My.ConnectExchangebtn.Add_Click({

$getsessions = Get-PSSession | Select-Object -Property State, Name
$isconnected = (@($getsessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
If ($isconnected -ne "True") {
Connect-ExchangeOnline
}


$Script:My.EXORoomlistsAddBtn.IsEnabled = $true
$Script:My.EXORoomlistsDelBtn.IsEnabled = $true
$Script:My.EXORoomlistsSearchTb1.IsEnabled = $true
$Script:My.EXORoomlistsSearchTb2.IsEnabled = $true
$Script:My.EXORoomlistsSearchBtn1.IsEnabled = $true
$Script:My.EXORoomlistsSearchBtn2.IsEnabled = $true

$Script:My.EXORoomlistView1.Items.Clear()
$Script:My.EXORoomlistView2.Items.Clear()
$Script:My.EXORoomlistView3.Items.Clear()

$arraylistgroups.Clear();
$arraylistrooms.Clear();
$arraylistmembers.Clear();


loadexoroomlistview

$Script:My.ConnectExchangebtn.IsEnabled = $false
$Script:My.CheckBox2.IsChecked = $true
})



#----------------------EXO-Roomlists--------------------

$Script:My.EXORoomlistsSearchBtn1.Add_Click({


$varEXORoomlistsSearchBtn1 = $Script:My.EXORoomlistsSearchTb1.Text.ToString()
if($varEXORoomlistsSearchBtn1 -eq "Search"){$varEXORoomlistsSearchBtn1=""}
$Script:My.EXORoomlistView1.Items.Clear()


$arraylistgroups.Clear();
$global:groups -match $varEXORoomlistsSearchBtn1 | Select-Object DisplayName, Id | ForEach-Object {$Script:My.EXORoomlistView1.Items.Add($_.DisplayName) |Out-Null
$arraylistgroups.Add($_.Id)

    }


    
})

$Script:My.EXORoomlistsSearchBtn2.Add_Click({


$varEXORoomlistsSearchBtn2 = $Script:My.EXORoomlistsSearchTb2.Text.ToString()
if($varEXORoomlistsSearchBtn2 -eq "Search"){$varEXORoomlistsSearchBtn2=""}

$Script:My.EXORoomlistView3.Items.Clear()

$arraylistrooms.Clear();
$global:rooms -match $varEXORoomlistsSearchBtn2 | ForEach-Object {$Script:My.EXORoomlistView3.Items.Add($_) |Out-Null
$arraylistrooms.Add($_)

    }
   
})

$Script:My.EXORoomlistsAddBtn.Add_Click({


$selected = $Script:My.EXORoomlistView1.SelectedIndex
$index1 = $arraylistgroups[$selected]


$selected = $Script:My.EXORoomlistView3.SelectedIndex
$index2 = $arraylistrooms[$selected]

Add-DistributionGroupMember -Id $index1 -Member "$index2"


})

$Script:My.EXORoomlistsDelBtn.Add_Click({


$selected = $Script:My.EXORoomlistView1.SelectedIndex
$index1 = $arraylistgroups[$selected]


$selected = $Script:My.EXORoomlistView2.SelectedIndex
$index2 = $arraylistmembers[$selected]

Remove-DistributionGroupMember -Id $index1 -Member "$index2"

loadexoroomlistview

})

#----------------------EXO-Roomlists--------------------

#--------------------EXO-Room-ListView------------------

$Script:My.EXORoomlistView1.Add_SelectionChanged({

$Script:My.EXORoomlistView2.Items.Clear()


$selected = $Script:My.EXORoomlistView1.SelectedIndex


$index = $arraylistgroups[$selected]


$members = Get-DistributionGroupMember -Id $index | select -ExpandProperty DisplayName | Sort-Object

$arraylistmembers.Clear();

$members  | ForEach-Object {$Script:My.EXORoomlistView2.Items.Add($_) |Out-Null
$arraylistmembers.Add($_)

    }

})

#--------------------EXO-Room-ListView------------------

#---------------------Functionblock---------------------
function loadexoroomlistview(){

$global:groups = Get-DistributionGroup -RecipientTypeDetails RoomList  
$global:rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox | select -ExpandProperty DisplayName


$global:groups = $global:groups | Sort-Object -Property "DisplayName"
$global:rooms = $global:rooms | Sort-Object -Property "DisplayName"

$Script:My.EXORoomlistView1.Items.Clear()

$arraylistgroups.Clear();
$global:groups | Select-Object DisplayName, Id | ForEach-Object {$Script:My.EXORoomlistView1.Items.Add($_.DisplayName) |Out-Null
$arraylistgroups.Add($_.Id)

    }

$Script:My.EXORoomlistView3.Items.Clear()

$arraylistrooms.Clear();
$global:rooms | ForEach-Object {$Script:My.EXORoomlistView3.Items.Add($_) |Out-Null
$arraylistrooms.Add($_)

    }


}

#---------------------Functionblock---------------------

#-------------------Clean-up-garbage--------------------

$Script:My.Window.Topmost = $true
$Script:My.Window.ShowDialog() | Out-Null


Remove-Variable -Name 'My' -Force
Set-StrictMode -Off
$ErrorActionPreference = $ErrorActionPreferenceBackup
Remove-Module -Name 'Microsoft.PowerShell.Management', 'Microsoft.PowerShell.Utility' -Force
#-------------------Clean-up-garbage--------------------
