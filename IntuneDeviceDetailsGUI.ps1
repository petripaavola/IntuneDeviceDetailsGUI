# Intune Device Details GUI ver 2.3
#
# This tool visualizes Intune device details
# Some of this information is not shown easily or at all in Intune web console
# - Application and Configuration Deployments include information what Azure AD Group was used for assignment and if filter was applied
# - Last signed in users are shown
# - Device JSON data inside Intune -> helps for example to build Azure AD Dynamic groups rules
# - This tool helps to understand why some App or Configuration Profile are applying to device (what Azure AD group and or filter applied)
#
# Tip: hover with mouse on top of different values. There is more info shown (for example PrimaryUser, Autopilot profile, OS/version)
#
# Examples:
#
# .\IntuneDeviceDetailsGUI.ps1
# .\IntuneDeviceDetailsGUI.ps1 -deviceName MyLoveMostPC
# .\IntuneDeviceDetailsGUI.ps1 -Id 2e6e1d5f-b18a-44c6-989e-9bbb1efafbff
# .\IntuneDeviceDetailsGUI.ps1 -IntuneDeviceId 2e6e1d5f-b18a-44c6-989e-9bbb1efafbff
# .\IntuneDeviceDetailsGUI.ps1 -serialNumber 1234567890
#
# # Pipe Intune objects to script from Powershell console
# Get-IntuneManagedDevice -Filter "devicename eq 'MyLoveMostPC'" | .\IntuneDeviceDetailsGUI.ps1
# 'MyLoveMostPC' | .\IntuneDeviceDetailsGUI.ps1
#
# # Or create Device Management UI with Out-GridView
# Get-IntuneManagedDevice | Out-GridView -OutputMode Single | .\IntuneDeviceDetailsGUI.ps1
#
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Devices for IT
# https://github.com/petripaavola/IntuneDeviceDetailsGUI
#
# Ver 2.3

[CmdletBinding(DefaultParameterSetName = 'deviceName')]
Param(
    [Parameter(Mandatory=$true,
				ParameterSetName = 'deviceName',
			    HelpMessage = 'Enter Intune deviceName',
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
	[Alias("displayName")]
    [String]$deviceName = $null,
	
    [Parameter(Mandatory=$true,
				ParameterSetName = 'id',
				HelpMessage = 'Enter Intune device ID',
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
	[ValidateScript({
        try {
            [System.Guid]::Parse($_) | Out-Null
            $true
        } catch {
            $false
        }
    })]
    [Alias("IntuneDeviceId")]
    [String]$id = $null,
	
	[Parameter(Mandatory=$true,
				ParameterSetName = 'serialNumber',
			    HelpMessage = 'Enter Intune device serialNumber',
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
	[Alias("serial")]	
    [String]$serialNumber = $null
)

$ScriptVersion = "ver 2.3"

$IntuneDeviceId = $id


#region XAML

#PUT XAML BELOW between the @" "@
$inputXML = @"
<Window x:Class="IntuneManagementGUI.IntuneDeviceDetails"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:IntuneManagementGUI"
        mc:Ignorable="d"
        Title="IntuneDeviceDetails" Height="800" MinHeight="730" Width="1055" MinWidth="1030" WindowStyle="ThreeDBorderWindow">
    <Window.TaskbarItemInfo>
        <TaskbarItemInfo/>
    </Window.TaskbarItemInfo>
    <Grid x:Name="IntuneDeviceDetails" Background="#FFE5EEFF">
        <Grid.ColumnDefinitions>
            <ColumnDefinition MinWidth="200" MaxWidth="200"/>
            <ColumnDefinition Width="*" MinWidth="400"/>
            <ColumnDefinition Width="*" MinWidth="200" MaxWidth="600"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="*" MinHeight="30" MaxHeight="30"/>
            <RowDefinition MinHeight="105" MaxHeight="105"/>
            <RowDefinition MinHeight="150"/>
            <RowDefinition MinHeight="150"/>
            <RowDefinition MinHeight="150"/>
            <RowDefinition Height="50" MinHeight="50" MaxHeight="50"/>
        </Grid.RowDefinitions>
        <Border x:Name="IntuneDeviceDetailsBorderTop" Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="3" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid x:Name="GridIntuneDeviceDetailsBorderTop" Grid.Row="0" ShowGridLines="False">
                <Button x:Name="Refresh_button" Content="Refresh" HorizontalAlignment="Right" VerticalAlignment="Center" Width="75" Margin="0,0,10,0"/>
                <TextBox x:Name="IntuneDeviceDetails_textBox_DeviceName" HorizontalAlignment="Left" Height="27" Margin="10,0,0,0" Text="DeviceName" VerticalAlignment="Top" Width="500" FontWeight="Bold" FontSize="20" Foreground="#FF004CFF" IsReadOnly="True"/>
            </Grid>
        </Border>
        <Border x:Name="IntuneDeviceDetailsBorderDeviceDetails" Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="3" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid x:Name="GridIntuneDeviceDetails" ShowGridLines="False" Margin="0.2,0.2,-0.4,-0.4">
                <Label x:Name="Manufacturer_label" Content="Manufacturer" HorizontalAlignment="Left" Margin="4,1,0,0" VerticalAlignment="Top" Height="24" FontWeight="Bold"/>
                <TextBox x:Name="Manufacturer_textBox" HorizontalAlignment="Left" Height="23" Margin="88,4,0,0" Text="Manufacturer" VerticalAlignment="Top" Width="165" IsReadOnly="True"/>
                <Label x:Name="Model_label" Content="Model" HorizontalAlignment="Left" Margin="4,25,0,0" VerticalAlignment="Top" Height="24" FontWeight="Bold"/>
                <TextBox x:Name="Model_textBox" HorizontalAlignment="Left" Height="23" Margin="88,28,0,0" Text="Model" VerticalAlignment="Top" Width="165" IsReadOnly="True"/>
                <Label x:Name="Serial_label" Content="Serial" HorizontalAlignment="Left" Margin="4,49,0,0" VerticalAlignment="Top" Height="24" FontWeight="Bold"/>
                <TextBox x:Name="Serial_textBox" HorizontalAlignment="Left" Height="23" Margin="88,52,0,0" Text="Serial" VerticalAlignment="Top" Width="165" IsReadOnly="True"/>
                <Label x:Name="WiFi_label" Content="WiFi MAC" HorizontalAlignment="Left" Margin="4,73,0,0" VerticalAlignment="Top" Height="24" FontWeight="Bold"/>
                <TextBox x:Name="WiFi_textBox" HorizontalAlignment="Left" Height="23" Margin="88,76,0,0" Text="WiFi" VerticalAlignment="Top" Width="165" IsReadOnly="True"/>
                <Label x:Name="OSVersion_label" Content="OS/Version" HorizontalAlignment="Left" Margin="262,1,0,0" VerticalAlignment="Top" Height="24" FontWeight="Bold"/>
                <TextBox x:Name="OSVersion_textBox" ToolTipService.ShowDuration="60000" HorizontalAlignment="Left" Height="23" Margin="346,4,0,0" Text="OS/Version" VerticalAlignment="Top" Width="165" IsReadOnly="True"/>
                <Label x:Name="Language_label" Content="Language" HorizontalAlignment="Left" Margin="262,25,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
                <TextBox x:Name="Language_textBox" HorizontalAlignment="Left" Height="23" Margin="346,28,0,0" Text="Language" VerticalAlignment="Top" Width="165" IsReadOnly="True"/>
                <Label x:Name="Storage_label" Content="Storage" HorizontalAlignment="Left" Margin="262,49,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
                <TextBox x:Name="Storage_textBox" HorizontalAlignment="Left" Height="23" Margin="346,52,52,0" Text="Storage" VerticalAlignment="Top" Width="165" IsReadOnly="True"/>
                <Label x:Name="EthernetMAC_label" Content="Ethernet MAC" HorizontalAlignment="Left" Margin="261,73,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
                <TextBox x:Name="EthernetMAC_textBox" HorizontalAlignment="Left" Height="23" Margin="346,76,52,0" Text="Ethernet MAC" VerticalAlignment="Top" Width="165" IsReadOnly="True"/>
                <Label x:Name="Compliance_label" Content="Compliance" HorizontalAlignment="Left" Margin="520,1,0,0" VerticalAlignment="Top" Height="24" FontWeight="Bold"/>
                <TextBox x:Name="Compliance_textBox" HorizontalAlignment="Left" Height="23" Margin="600,4,0,0" Text="Compliance" VerticalAlignment="Top" Width="177" IsReadOnly="True"/>
                <Label x:Name="isEncrypted_label" Content="isEncrypted" HorizontalAlignment="Left" Margin="520,25,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
                <TextBox x:Name="isEncrypted_textBox" HorizontalAlignment="Left" Height="23" Margin="600,28,0,0" Text="isEncrypted" VerticalAlignment="Top" Width="177" IsReadOnly="True"/>
                <Label x:Name="lastSync_label" Content="Last Sync" HorizontalAlignment="Left" Margin="520,49,0,0" VerticalAlignment="Top" Height="24" FontWeight="Bold"/>
                <TextBox x:Name="lastSync_textBox" ToolTipService.ShowDuration="60000" HorizontalAlignment="Left" Height="23" Margin="600,52,0,0" Text="Last Sync" VerticalAlignment="Top" Width="177" IsReadOnly="True"/>
                <Label x:Name="primaryUser_label" Content="Primary User" HorizontalAlignment="Left" Margin="520,73,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
                <TextBox x:Name="primaryUser_textBox" ToolTipService.ShowDuration="60000" HorizontalAlignment="Left" Height="23" Margin="600,76,0,0" Text="Primary User" VerticalAlignment="Top" Width="177" IsReadOnly="True"/>
                <Label x:Name="AutopilotGroup_label" Content="Windows Autopilot" HorizontalAlignment="Left" Margin="790,1,0,0" VerticalAlignment="Top" Height="24" FontWeight="Bold"/>
                <Label x:Name="AutopilotEnrolled_label" Content="enrolled" HorizontalAlignment="Left" Margin="790,25,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
                <TextBox x:Name="AutopilotEnrolled_textBox" ToolTipService.ShowDuration="60000" HorizontalAlignment="Left" Height="23" Margin="858,28,0,0" Text="" VerticalAlignment="Top" Width="177" IsReadOnly="True"/>
                <Label x:Name="AutopilotGroupTag_label" Content="GroupTag" HorizontalAlignment="Left" Margin="790,49,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
                <Label x:Name="AutopilotProfile_label" Content="Profile" HorizontalAlignment="Left" Margin="790,73,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
                <TextBox x:Name="AutopilotGroupTag_textBox" ToolTipService.ShowDuration="60000" HorizontalAlignment="Left" Height="23" Margin="858,52,0,0" Text="" VerticalAlignment="Top" Width="177" IsReadOnly="True"/>
                <TextBox x:Name="AutopilotProfile_textBox" ToolTipService.ShowDuration="60000" HorizontalAlignment="Left" Height="23" Margin="858,76,0,0" Text="" VerticalAlignment="Top" Width="177" IsReadOnly="True"/>
            </Grid>
        </Border>
        <Border x:Name="IntuneDeviceDetailsBorderSignInUsers" Grid.Row="2" Grid.Column="0" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="25" MinHeight="25"/>
                    <RowDefinition Height="*" MinHeight="300"/>
                </Grid.RowDefinitions>
                <Label x:Name="IntuneDeviceDetails_RecentCheckins_label" Grid.Row="0" Content="Recent check-ins" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
                <TextBox x:Name="IntuneDeviceDetails_RecentCheckins_textBox" Grid.Row="1" Margin="5,5,5,5" TextWrapping="Wrap" Text="" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
            </Grid>
        </Border>
        <Border x:Name="IntuneDeviceDetailsBorderGroupMemberships" Grid.Row="2" Grid.Column="1" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="25" MinHeight="25"/>
                    <RowDefinition Height="*" MinHeight="300"/>
                </Grid.RowDefinitions>
                <Label x:Name="IntuneDeviceDetails_GroupMemberships_label" Grid.Row="0" Content="Device Group Membership" HorizontalAlignment="Left" Margin="8,0,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
                <TextBox x:Name="IntuneDeviceDetails_GroupMemberships_textBox" Grid.Row="1" Margin="5,5,5,5" TextWrapping="Wrap" Text="" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
            </Grid>
        </Border>
        <Border x:Name="IntuneDeviceDetailsBorderXAML" Grid.Row="2" Grid.Column="2" Grid.RowSpan="3" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="25" MinHeight="25"/>
                    <RowDefinition Height="*" MinHeight="300"/>
                </Grid.RowDefinitions>
                <Label x:Name="IntuneDeviceDetails_json_label" Grid.Row="0" Content="Device JSON data" HorizontalAlignment="Left" Margin="8,0,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
                <TextBox x:Name="IntuneDeviceDetails_json_textBox" Grid.Row="1" TextWrapping="Wrap" Text="" Padding="5" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Margin="5" IsReadOnly="True" />
            </Grid>
        </Border>
        <Border x:Name="IntuneDeviceDetailsBorderApplications" Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="25" MinHeight="25"/>
                    <RowDefinition Height="*" MinHeight="150"/>
                </Grid.RowDefinitions>
                <Label x:Name="IntuneDeviceDetails_ApplicationAssignments_label" Grid.Row="0" Content="Application Assignments" Height="27" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
                <!-- <TextBox x:Name="IntuneDeviceDetails_ApplicationAssignments_textBox" Grid.Row="1" Margin="5,5,5,5" Text="" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/> -->
                <ListView x:Name="listView_ApplicationAssignments" Grid.Row="1" Margin="5,5,5,5" IsManipulationEnabled="True">
                    <!-- This makes our colored cells to fill whole cell background, not just text background -->
                    <ListView.ItemContainerStyle>
                        <Style TargetType="ListViewItem">
                            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                        </Style>
                    </ListView.ItemContainerStyle>
                    <ListView.ContextMenu>
                        <ContextMenu IsTextSearchEnabled="True">
                            <!-- <MenuItem x:Name = 'AutopilotTAB_ListView_CopyToClipBoardJSON_Menu' Header = 'Copy to clipboard JSON'/> -->
                        </ContextMenu>
                    </ListView.ContextMenu>
                    <ListView.View>
                        <GridView>
                            <!-- <GridViewColumn Width="60" Header="context" DisplayMemberBinding="{Binding 'context'}"/> -->
                            <GridViewColumn Width="70" Header="context">
                                <GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
                                            <Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
                                                        <DataTrigger Binding="{Binding context}" Value="_unknown">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
                                                        <DataTrigger Binding="{Binding context}" Value="_Device/User">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </Grid.Style>
                                            <TextBlock HorizontalAlignment="Center" Text="{Binding context}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
                            <GridViewColumn Width="175" Header="Application type" DisplayMemberBinding="{Binding 'odatatype'}"/>
                            <!-- <GridViewColumn Width="160" Header="displayName" DisplayMemberBinding="{Binding displayName}"/> -->
                            <GridViewColumn Width="160" Header="displayName">
                                <GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock FontWeight="Bold" Text="{Binding displayName}" >
                                        </TextBlock>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
                            <GridViewColumn Width="70" Header="version" DisplayMemberBinding="{Binding version}"/>
                            <GridViewColumn Width="95" Header="intent" DisplayMemberBinding="{Binding assignmentIntent}"/>
                            <GridViewColumn Width="90" Header="IncludeExclude" DisplayMemberBinding="{Binding IncludeExclude}"/>
                            <!-- <GridViewColumn Width="140" Header="installState" DisplayMemberBinding="{Binding installState}"/> -->
                            <GridViewColumn Width="130" Header="installState">
                                <GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
                                            <Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
                                                        <DataTrigger Binding="{Binding installState}" Value="failed">
                                                            <Setter Property="Background" Value="#FF6347"/>
                                                        </DataTrigger>
                                                        <DataTrigger Binding="{Binding installState}" Value="installed">
                                                            <Setter Property="Background" Value="#7FFF00"/>
                                                        </DataTrigger>
                                                        <!-- <DataTrigger Binding="{Binding installState}" Value="unknown">
                                                            <Setter Property="Background" Value="#FF6347"/>
                                                        </DataTrigger> -->
                                                    </Style.Triggers>
                                                </Style>
                                            </Grid.Style>
                                            <TextBlock HorizontalAlignment="Center" Text="{Binding installState}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
                            <GridViewColumn Width="250" Header="assignmentGroup" DisplayMemberBinding="{Binding assignmentGroup}"/>
							<GridViewColumn Width="250" Header="Filter" DisplayMemberBinding="{Binding filter}"/>
							<GridViewColumn Width="80" Header="FilterMode" DisplayMemberBinding="{Binding filterMode}"/>
                            <!-- <GridViewColumn Width="140" Header="modified" DisplayMemberBinding="{Binding modified}"/> -->
                        </GridView>
                    </ListView.View>
                </ListView>
            </Grid>
        </Border>
        <Border x:Name="IntuneDeviceDetailsBorderConfigurations" Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="25" MinHeight="25"/>
                    <RowDefinition Height="*" MinHeight="150"/>
                </Grid.RowDefinitions>
                <Label x:Name="IntuneDeviceDetails_ConfigurationsAssignments_label" Grid.Row="0" Content="Configurations Assignments" Height="27" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
                <!-- <TextBox x:Name="IntuneDeviceDetails_ConfigurationsAssignments_textBox" Grid.Row="1" Margin="5,5,5,5" TextWrapping="Wrap" Text="" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/> -->
                <ListView x:Name="listView_ConfigurationsAssignments" Grid.Row="1" Margin="5,5,5,5" IsManipulationEnabled="True">
                    <!-- This makes our colored cells to fill whole cell background, not just text background -->
                    <ListView.ItemContainerStyle>
                        <Style TargetType="ListViewItem">
                            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                        </Style>
                    </ListView.ItemContainerStyle>
                    <ListView.ContextMenu>
                        <ContextMenu IsTextSearchEnabled="True">
                            <!-- <MenuItem x:Name = 'AutopilotTAB_ListView_CopyToClipBoardJSON_Menu' Header = 'Copy to clipboard JSON'/> -->
                        </ContextMenu>
                    </ListView.ContextMenu>
                    <ListView.View>
                        <GridView>
                            <!-- <GridViewColumn Width="60" Header="context" DisplayMemberBinding="{Binding 'context'}"/> -->
                            <GridViewColumn Width="70" Header="context">
                                <GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
                                            <Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
                                                        <DataTrigger Binding="{Binding context}" Value="_unknown">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
                                                        <DataTrigger Binding="{Binding context}" Value="_Device/User">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </Grid.Style>
                                            <TextBlock HorizontalAlignment="Center" Text="{Binding context}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
                            <GridViewColumn Width="175" Header="Configuration type" DisplayMemberBinding="{Binding 'odatatype'}"/>
                            <!-- <GridViewColumn Width="160" Header="displayName" DisplayMemberBinding="{Binding displayName}"/> -->
                            <GridViewColumn Width="160" Header="displayName">
                                <GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock FontWeight="Bold" Text="{Binding displayName}" >
                                        </TextBlock>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
                            <GridViewColumn Width="175" Header="userPrincipalName" DisplayMemberBinding="{Binding 'userPrincipalName'}"/>
                            <GridViewColumn Width="80" Header="IncludeExclude" DisplayMemberBinding="{Binding IncludeExclude}"/>
                            <GridViewColumn Width="85" Header="state">
                                <GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
                                            <Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
                                                        <!-- <DataTrigger Binding="{Binding state}" Value="Not applicable">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger> -->
                                                        <DataTrigger Binding="{Binding state}" Value="Succeeded">
                                                            <Setter Property="Background" Value="#7FFF00"/>
                                                        </DataTrigger>
														<DataTrigger Binding="{Binding state}" Value="Conflict">
                                                            <Setter Property="Background" Value="#FF6347"/>
                                                        </DataTrigger>
														<DataTrigger Binding="{Binding state}" Value="Error">
                                                            <Setter Property="Background" Value="#FF6347"/>
                                                        </DataTrigger>
                                                        <!-- <DataTrigger Binding="{Binding state}" Value="unknown">
                                                        <Setter Property="Background" Value="#FF6347"/>
                                                    </DataTrigger> -->
                                                    </Style.Triggers>
                                                </Style>
                                            </Grid.Style>
                                            <TextBlock HorizontalAlignment="Center" Text="{Binding state}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
                            <GridViewColumn Width="250" Header="assignmentGroup" DisplayMemberBinding="{Binding assignmentGroup}"/>
							<GridViewColumn Width="250" Header="Filter" DisplayMemberBinding="{Binding filter}"/>
							<GridViewColumn Width="80" Header="FilterMode" DisplayMemberBinding="{Binding filterMode}"/>
                            <!-- <GridViewColumn Width="140" Header="modified" DisplayMemberBinding="{Binding modified}"/> -->
                        </GridView>
                    </ListView.View>
                </ListView>
            </Grid>
        </Border>
        <Border x:Name="IntuneDeviceDetailsBorderBottom" Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="3" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid x:Name="IntuneDeviceDetailsGridBottom">
                <Image x:Name="bottomRightLogoimage" HorizontalAlignment="Right" Height="46" Margin="0,0,9.8,0" VerticalAlignment="Top" Width="133"/>
                <TextBox x:Name="bottom_textBox" HorizontalAlignment="Left" Height="26" Margin="10,10,0,0" Text="" VerticalAlignment="Top" Width="609" FontSize="16" IsReadOnly="True"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

#endregion XAML


#region Load XAML
$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
#
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
#
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {$Form = [Windows.Markup.XamlReader]::Load( $reader )}
catch {Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
#
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
#
$xaml.SelectNodes("//*[@Name]") | % {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
#
Function Get-FormVariables {
    if ($global:ReadmeDisplay -ne $true) {Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow; $global:ReadmeDisplay = $true}
    write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
    get-variable WPF*
}
# Debug print all form variables to console
# Comment this out when running in production
#Get-FormVariables
#Get-FormVariables | Select-Object -Expandproperty Name

#endregion Load XAML

##################################################################################################

#region General functions

function DecodeBase64Image {
        param (
        [Parameter(Mandatory = $true)]
        [String]$ImageBase64
    )
    # Parameter help description
    $ObjBitmapImage = New-Object System.Windows.Media.Imaging.BitmapImage #Provides a specialized BitmapSource that is optimized for loading images using Extensible Application Markup Language (XAML).
    $ObjBitmapImage.BeginInit() #Signals the start of the BitmapImage initialization.
    $ObjBitmapImage.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($ImageBase64) #Creates a stream whose backing store is memory.
    $ObjBitmapImage.EndInit() #Signals the end of the BitmapImage initialization.
    $ObjBitmapImage.Freeze() #Makes the current object unmodifiable and sets its IsFrozen property to true.
    $ObjBitmapImage
}

function Convert-Base64ToFile {
    Param(
        [String]$base64,
        $filepath
    )

    $bytes = [Convert]::FromBase64String($base64)
    [IO.File]::WriteAllBytes($filepath, $bytes)
    $Success = $?

    return $Success
}

function Convert-UTCtoLocal {
    param(
        [parameter(Mandatory=$true)]
        [String] $UTCTime
    )
    $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
    $TimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
    $LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TimeZone)
    return $LocalTime
}


##################################################################################################
#endregion General functions

##################################################################################################

function Invoke-MSGraphGetRequestWithMSGraphAllPages {
    param (
        [Parameter(Mandatory = $true)]
        [String]$url
    )

    $MSGraphRequest = $null
    $AllMSGraphRequest = $null

    try {
        $MSGraphRequest = Invoke-MSGraphRequest -Url $url -HttpMethod 'GET'
        $Success = $?

        if($Success) {
            
            # This does not work because we won't catch this if there is Value-attribute which is null
            #if ($MSGraphRequest.Value) {

            # Test if object has attribute named Value (whether value is null or not)
            if((Get-Member -inputobject $MSGraphRequest -name 'Value' -Membertype Properties) -and (Get-Member -inputobject $MSGraphRequest -name '@odata.context' -Membertype Properties)) {
                # Value property exists
                $returnObject = $MSGraphRequest.Value
            } else {
                # Sometimes we get results without Value-attribute (eg. getting user details)
                # We will return all we got as is
                $returnObject = $MSGraphRequest
            }
        } else {
            # Invoke-MSGraphRequest failed so we return false
            return $null
        }

        # Check if we have value starting https:// in attribute @odate.nextLink
        # If we have nextLink then we get GraphAllPages
        if ($MSGraphRequest.'@odata.nextLink' -like 'https://*') {

            # Get AllMSGraph pages
            # This is also workaround to get objects without assigning them from .Value attribute
            $AllMSGraphRequest = Get-MSGraphAllPages -SearchResult $MSGraphRequest
            $Success = $?

            if($Success) {
                $returnObject = $AllMSGraphRequest
            } else {
                # Getting Get-MSGraphAllPages failed
                return $null
            }
        }

        return $returnObject

    } catch {
        Write-Error "There was error with MSGraphRequest with url $url!"
        return $null
    }
}

function Get-ApplicationsWithAssignments {

    # Check if Apps have changed in Intune after last cached file was loaded
    # We try to get Apps changed after last cache file modified date

    $AppsWithAssignmentsFilePath = "$PSScriptRoot\cache\$TenantId\AllApplicationsWithAssignments.json"

    # Check if AllApplicationsWithAssignments.json file exists
    if(Test-Path "$AppsWithAssignmentsFilePath") {

        $FileDetails = Get-ChildItem "$AppsWithAssignmentsFilePath"

        # Get timestamp for file AllApplicationsWithAssignments.json
        # We use uformat because Culture can otherwise change separator for time format (change : -> .)
        # Get-Date -uformat %G-%m-%dT%H:%M:%S.0000000Z
        $AppsFileLastWriteTimeUtc = Get-Date $FileDetails.LastWriteTimeUtc -uformat %G-%m-%dT%H:%M:%S.000Z

        # Get MobileApps modified after cache file datetime
        #https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$filter=lastModifiedDateTime%20gt%202019-12-31T00:00:00.000Z&$expand=assignments
        $url = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=lastModifiedDateTime%20gt%20$AppsFileLastWriteTimeUtc&`$expand=assignments&`$top=1000"

        $AllAppsChangedAfterCacheFileDate = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

        if ($AllAppsChangedAfterCacheFileDate) {
            # We found new/changed Apps which we don't have in our cache file so we need to download Apps

            # Future TODO: get changed Apps and migrate that to existing cache file
            # and Always force update Apps after 7 days old cache file

            # For now we don't actually do anything here because next phase will download All Apps and update cache file
        } else {
            # We found no changed Apps so our cache file is still valid
            # We can use cached file
            
            $AppsWithAssignments = Get-Content "$AppsWithAssignmentsFilePath" | ConvertFrom-Json
            return $AppsWithAssignments
        }
    } 

	# If we end up here then file either does not exist or we need to update existing cached file
	$url = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$expand=assignments&`$top=1000&_=1577625591870"

	$AllAppsWithAssignments = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

	if($AllAppsWithAssignments) {
		# Save to local cache
		# Use -Depth 4 to get all Assignment information also !!!
		$AllAppsWithAssignments | ConvertTo-Json -Depth 4 | Out-File "$AppsWithAssignmentsFilePath" -Force
		
		# Load Application information from cached file always
		$AppsWithAssignments = Get-Content "$AppsWithAssignmentsFilePath" | ConvertFrom-Json
		return $AppsWithAssignments
	} else {
		return $false
	}
}

function Objectify_JSON_Schema_and_Data_To_PowershellObjects {
	Param(
		[Parameter(Mandatory=$true,
			ValueFromPipeline=$true,
			ValueFromPipelineByPropertyName=$true, 
			Position=0)]
		$Json
	)

	$JsonSchema = $Json.Schema
	$JsonValues = $Json.Values

	# Create empty arrayList
	# ArrayList is quicker if we have a huge data set
	# because using array and += always creates new array with added array value/object
	$JsonObjectArrayList = New-Object -TypeName "System.Collections.ArrayList"

	# Convert json data to Powershell objects in $JsonObjectArrayList
	foreach($Value in $JsonValues) {
		# We use this counter to get property name value from Schema array
		$i=0

		# Add values to HashTable which we use to create custom Powershell object later
		$ValuesHashTable = @{}

		foreach($ValueEntry in $Value) {
			# Create variables
			$PropertyName = $JsonSchema[$i].Column
			$ValuePropertyType = $JsonSchema[$i].PropertyType
			$PropertyValue = $ValueEntry -as $ValuePropertyType
			
			# Add hashtable entry
			$ValuesHashTable.add($PropertyName, $PropertyValue)

			# Create Powershell custom object from hashtable
			$CustomObject = new-object psobject -Property $ValuesHashTable
			
			$i++
		}

		# Add custom Powershell object to ArrayList
		$JsonObjectArrayList.Add($CustomObject) | Out-Null
	}

	return $JsonObjectArrayList
}


function Download-IntuneFilters {
	try {
		$url = 'https://graph.microsoft.com/beta/deviceManagement/assignmentFilters'
		$MSGraphRequest = Invoke-MSGraphRequest -Url $url -HttpMethod 'GET'
		$Success = $?

		if (-not ($Success)) {
			Write-Error "Error downloading Intune filters information"
			Write-Error "Script will exit..."
			Pause
			Exit 1
		}

		$AllIntuneFilters = Get-MSGraphAllPages -SearchResult $MSGraphRequest
		
		Write-Verbose "Found $($AllIntuneFilters.Count) Intune filters"

		return $AllIntuneFilters

    } catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download Intune filters information"
        Write-Error "Script will exit..."
        Pause
        Exit 1
    }
}


function Download-IntuneConfigurationProfiles {

	# Intune Policies to download including Assignment information
	# Note that Endpint Security configurations (intents) do not expand assignments!!!

<#
	# Without selection filtering -> gets all data
	$urls = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?expand=assignments',`
			'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?expand=assignments',`
			'https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?expand=assignments',`
			'https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations?expand=assignments',`
			'https://graph.microsoft.com/beta/deviceManagement/intents'
#>

	# Using Select to filter only necessary data -> smaller downloads
	$urls = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?expand=assignments&$select=id,description,createdDateTime,lastModifiedDateTime,name,assignments',`
			'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?expand=assignments&$select=id,description,createdDateTime,lastModifiedDateTime,displayname,assignments',`
			'https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?expand=assignments&$select=id,description,createdDateTime,lastModifiedDateTime,displayname,assignments',`
			'https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations?expand=assignments&$select=id,description,createdDateTime,lastModifiedDateTime,displayname,assignments',`
			'https://graph.microsoft.com/beta/deviceManagement/intents'


	# CustomObject ArrayList which is returned
	$ConfigurationsArray = @()

	foreach($url in $urls) {
		try {
			$MSGraphRequest = Invoke-MSGraphRequest -Url $url -HttpMethod 'GET'
			$Success = $?

			if (-not ($Success)) {
				Write-Error "Error downloading Intune configurations: $url"
				Write-Error "Script will exit..."
				Pause
				Exit 1
			}

			$MSGraphRequestAllPages = Get-MSGraphAllPages -SearchResult $MSGraphRequest

		} catch {
			Write-Error "$($_.Exception.GetType().FullName)"
			Write-Error "$($_.Exception.Message)"
			Write-Error Write-Error "Error downloading Intune configurations (MSGraphAllPages): $url"
			Write-Error "Script will exit..."
			Pause
			Exit 1
		}
		$ConfigurationsArray += $MSGraphRequestAllPages
	}
	
	return $ConfigurationsArray
}


###################################################################################################################

function Get-DeviceInformation {

    # Update Graph API schema to beta to get Win32LobApps and possible other new features also
    Update-MSGraphEnvironment -SchemaVersion 'beta'
    $Success = $?

    if (-not $Success) {
        Write-Host "Failed to update MSGraph Environment schema to Beta!" -ForegroundColor Red
		Write-Host "Make sure you have installed Intune Powershell module"
		Write-Host "You can install Intune module with command: Install-Module -Name Microsoft.Graph.Intune" -ForegroundColor Yellow
		Write-Host "More information: https://github.com/microsoft/Intune-PowerShell-SDK"
        Exit 1
    }

    $MSGraphEnvironment = Connect-MSGraph
    $Success = $?

    if (-not $Success) {
        Write-Error "Could not connect to MSGraph!"
        Exit 1
    }

    $TenantId = $MSGraphEnvironment.tenantId

    # Create tenant specific cache folder if not exist
    # This is used for image caching used in GUI and Apps logos
    if(-not (Test-Path "$PSScriptRoot\cache\$TenantId")) {
        New-Item "$PSScriptRoot\cache\$TenantId" -Itemtype Directory
    }


	# If deviceName was specified as parameter then we first try to find device
	if($deviceName) {
		Write-Host "Getting Intune device information for $deviceName"
		$IntuneManagedDevice = Get-IntuneManagedDevice -Filter "devicename eq '$deviceName'"
		$Success = $?

		if (-not $Success) {
			Write-Host "Error getting intune device information!" -ForegroundColor Red
			Exit 1
		}
		
		if($IntuneManagedDevice) {
			if($IntuneManagedDevice -is [array]) {
				Write-Host "Warning: found $($IntuneManagedDevice.Count) devices from Intune with same deviceName" -ForegroundColor Yellow
				Write-Host "Getting device which has newest enrollmentDate"
				$IntuneManagedDevice = $IntuneManagedDevice | Sort-Object -Property enrolledDateTime -Descending | Select-Object -First 1

				$IntuneDeviceId = $IntuneManagedDevice.id
			} else {
				Write-Host "Found 1 device with Intune deviceId $($IntuneManagedDevice.id)"
				$IntuneDeviceId = $IntuneManagedDevice.id
			}
		} else {
			Write-Host "Did not find any devices with deviceName $deviceName" -ForegroundColor Red
		}
	}

	# If Intune Device ID was specified as parameter
	if($IntuneDeviceId) {
		Write-Host "Getting Intune device information for Intune deviceId: $IntuneDeviceId"

		$IntuneManagedDevice = Get-IntuneManagedDevice -managedDeviceId $IntuneDeviceId
		$Success = $?

		if (-not $Success) {
			Write-Host "Error getting intune device information!" -ForegroundColor Red
			Exit 1
		}

	}

	# if Intune serialNumber was specified as parameter
	if($serialNumber) {
		Write-Host "Getting Intune device information for serialNumber $serialNumber"
		$IntuneManagedDevice = Get-IntuneManagedDevice -Filter "serialNumber eq '$serialNumber'"
		$Success = $?

		if (-not $Success) {
			Write-Host "Error getting intune device information!" -ForegroundColor Red
			Exit 1
		}

		if($IntuneManagedDevice) {
			if($IntuneManagedDevice -is [array]) {
				Write-Host "Warning: found $($IntuneManagedDevice.Count) devices from Intune with same serialNumber" -ForegroundColor Yellow
				Write-Host "Getting device which has newest enrollmentDate"
				$IntuneManagedDevice = $IntuneManagedDevice | Sort-Object -Property enrolledDateTime -Descending | Select-Object -First 1

				$IntuneDeviceId = $IntuneManagedDevice.id
			} else {
				Write-Host "Found 1 device with Intune serialNumber $serialNumber ($($IntuneManagedDevice.deviceName) - $($IntuneManagedDevice.id))"
				$IntuneDeviceId = $IntuneManagedDevice.id
			}
		} else {
			Write-Host "Did not find any devices with serialNumber $serialNumber" -ForegroundColor Red
		}		
	}


    # If we didn't find device
    if(-not ($IntuneManagedDevice)) {
        $WPFIntuneDeviceDetails_textBox_DeviceName.Text = "Could not find device"
        $WPFIntuneDeviceDetails_textBox_DeviceName.Foreground = "red"
        return
    }


	Write-Host "Download Intune filters"
	$AllIntuneFilters = Download-IntuneFilters
	Write-Host "Found $($AllIntuneFilters.Count) filters"

	Write-Host "Download Intune configuration profiles"
	$IntuneConfigurationProfilesWithAssignments = Download-IntuneConfigurationProfiles
	Write-Host "Found $($IntuneConfigurationProfilesWithAssignments.Count) configuration profiles"

	# DEBUG
	# Copy json data to clipboard so you can paste json to any text editor
	#$IntuneConfigurationProfilesWithAssignments | ConvertTo-Json -Depth 6 | Set-Clipboard


    $WPFIntuneDeviceDetails_textBox_DeviceName.Text = $IntuneManagedDevice.DeviceName
    $WPFManufacturer_textBox.Text = $IntuneManagedDevice.Manufacturer
    $WPFModel_textBox.Text = $IntuneManagedDevice.Model
    $WPFSerial_textBox.Text = $IntuneManagedDevice.serialNumber
    $WPFWiFi_textBox.Text = $IntuneManagedDevice.wiFiMacAddress

    # These does not seem to work when getting only 1 device
    # Getting all devices results ok values
    # Workaround is to use hardwareInformation attribute
    #$IntuneManagedDevice.totalStorageSpaceInBytes
    #$IntuneManagedDevice.freeStorageSpaceInBytes

    # We get this value from 
    #$WPFEthernetMAC_textBox.Text = $IntuneManagedDevice.ethernetMacAddress

    # Get Additional Device information by specifying attributes
    # Check we have valid GUID
    if([System.Guid]::Parse($IntuneDeviceId)) {
        # Get additional device information
        $url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($IntuneDeviceId)?`$select=id,hardwareinformation,activationLockBypassCode,iccid,udid,roleScopeTagIds,ethernetMacAddress,processorArchitecture"
        $AdditionalDeviceInformation = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
    }

    if($AdditionalDeviceInformation) {

        # Windows SKU
        $operatingSystemEdition = ($AdditionalDeviceInformation.hardwareinformation).operatingSystemEdition

        if($operatingSystemEdition -eq 'Enterprise') { $operatingSystemEdition = 'Ent' }
        if($operatingSystemEdition -eq 'Education') { $operatingSystemEdition = 'Edu' }

        # EthernetMacAddress
        $WPFEthernetMAC_textBox.Text = $AdditionalDeviceInformation.ethernetMacAddress

        # OS Language
        $WPFLanguage_textBox.Text = ($AdditionalDeviceInformation.hardwareinformation).operatingSystemLanguage

        # free/total space
        $totalStorageGB = [math]::round(($AdditionalDeviceInformation.hardwareinformation).totalStorageSpace/1GB, 0)
        $freeStorageGB = [math]::round(($AdditionalDeviceInformation.hardwareinformation).freeStorageSpace/1GB, 0)
        $WPFStorage_textBox.Text = "$($freeStorageGB)GB `/ $($totalStorageGB)GB"

        # For example Android device may not show total and free space at all so we don't color on those devices
        if($totalStorageGB -gt 0) {
            if ($freeStorageGB -lt 10) {
                # Red
                $WPFStorage_textBox.Background = '#FF6347'
                $WPFStorage_textBox.Foreground = '#000000'
            } elseif (($freeStorageGB -ge 10) -and ($freeStorageGB -lt 15)) {
                # Yellow
                $WPFStorage_textBox.Background = 'yellow'
                $WPFStorage_textBox.Foreground = '#000000'
            } elseif ($freeStorageGB -ge 15) {
                # Green
                $WPFStorage_textBox.Background = '#7FFF00'
                $WPFStorage_textBox.Foreground = '#000000'
            }
        }
    }

    $CurrentDate = Get-Date -Format 'yyyy-MM-dd'
    [String]$WindowsSupportEndsInDays = $null
    [String]$Version = $null

    if($IntuneManagedDevice.operatingSystem -eq 'Windows') {
        switch -wildcard ($IntuneManagedDevice.osVersion) {
            '10.0.10240.*' {   $Version = '10 1507'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2017-05-09 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2017-05-09 | Select-Object -ExpandProperty Days
                        }
                    }
            '10.0.10586.*' {   $Version = '10 1511'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2017-10-10 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2017-10-10 | Select-Object -ExpandProperty Days
                        }
                    }
            '10.0.14393.*' {   $Version = '10 1607'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2019-04-09 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2018-04-10 | Select-Object -ExpandProperty Days
                        }
                    }
            '10.0.15063.*' {   $Version = '10 1703'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2019-10-08 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2018-10-09 | Select-Object -ExpandProperty Days
                        }
                    }
            '10.0.16299.*' {   $Version = '10 1709'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2020-10-13 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2019-05-09 | Select-Object -ExpandProperty Days
                        }
                    }
            '10.0.17134.*' {   $Version = '10 1803'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2020-11-10 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2019-11-12 | Select-Object -ExpandProperty Days
                        }
                    }
            '10.0.17763.*' {   $Version = '10 1809'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {            
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2021-05-11 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2020-05-12 | Select-Object -ExpandProperty Days
                        }
                    }
            '10.0.18362.*' {   $Version = '10 1903'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {            
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2020-12-08 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2020-12-08 | Select-Object -ExpandProperty Days
                        }
                    }
            '10.0.18363.*' {   $Version = '10 1909'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2022-05-10 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2021-05-11 | Select-Object -ExpandProperty Days
                        }
                    }
			'10.0.19041.*' {   $Version = '10 2004'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2021-12-14 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2021-12-14 | Select-Object -ExpandProperty Days
                        }
                    }
			'10.0.19042.*' {   $Version = '10 20H2'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2023-05-09 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2022-05-10 | Select-Object -ExpandProperty Days
                        }
                    }
			'10.0.19043.*' {   $Version = '10 21H1'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2022-12-13 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2022-12-13 | Select-Object -ExpandProperty Days
                        }
                    }
			'10.0.19044.*' {   $Version = '10 21H2'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2024-06-11 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2023-06-13 | Select-Object -ExpandProperty Days
                        }
                    }
			'10.0.22000.*' {   $Version = '11 21H2'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2024-10-08 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2023-10-10 | Select-Object -ExpandProperty Days
                        }
                    }
            Default {
                        $Version = $IntuneManagedDevice.operatingSystem
                    }
        }

        if([double]$WindowsSupportEndsInDays -lt 0) {
            # Windows 10 support already ended
            $OSVersionToolTip = "Windows 10 $($IntuneManagedDevice.osVersion)`n`nWindows 10 Enterprise and Education support for this version has already ended $WindowsSupportEndsInDays days ago!`n`nUpdate device immediately!"
            
            # Red background for OSVersion textbox
            $WPFOSVersion_textBox.Background = '#FF6347'
            $WPFOSVersion_textBox.Foreground = '#000000'

        } elseif (([double]$WindowsSupportEndsInDays -ge 0) -and ([double]$WindowsSupportEndsInDays -le 30)) {
            # Windows 10 support is ending in 30 days
            $OSVersionToolTip = "Windows 10 $($IntuneManagedDevice.osVersion)`n`nWindows 10 Enterprise and Education support for this version is ending in $WindowsSupportEndsInDays days.`n`nSchedule Windows upgrade for this device."
            
            # Yellow background for OSVersion textbox
            $WPFOSVersion_textBox.Background = 'yellow'
            $WPFOSVersion_textBox.Foreground = '#000000'
            
        } elseif([double]$WindowsSupportEndsInDays -gt 30) {
            # Windows 10 has support over 30 days
            $OSVersionToolTip = "Windows 10 $($IntuneManagedDevice.osVersion)`n`nWindows 10 Enterprise and Education support for this version will end in $WindowsSupportEndsInDays days."

            # Green background for OSVersion textbox
            $WPFOSVersion_textBox.Background = '#7FFF00'
            $WPFOSVersion_textBox.Foreground = '#000000'
        }
        $WPFOSVersion_textBox.Tooltip = $OSVersionToolTip

    } else {
        $Version = $IntuneManagedDevice.osVersion
    }
    $WPFOSVersion_textBox.Text = "$($IntuneManagedDevice.operatingSystem) $Version $operatingSystemEdition"

    $WPFCompliance_textBox.Text = $IntuneManagedDevice.complianceState
    if($IntuneManagedDevice.complianceState -eq 'compliant') {
        $WPFCompliance_textBox.Background = '#7FFF00'
        $WPFCompliance_textBox.Foreground = '#000000'
    }

    if($IntuneManagedDevice.complianceState -eq 'noncompliant') {
        $WPFCompliance_textBox.Background = '#FF6347'
        $WPFCompliance_textBox.Foreground = '#000000'
    }

    if($IntuneManagedDevice.complianceState -eq 'unknown') {
        $WPFCompliance_textBox.Background = 'yellow'
        $WPFCompliance_textBox.Foreground = '#000000'
    }

    $WPFisEncrypted_textBox.Text = $IntuneManagedDevice.isEncrypted
    if($IntuneManagedDevice.isEncrypted -eq 'True') {
        $WPFisEncrypted_textBox.Background = '#7FFF00'
        $WPFisEncrypted_textBox.Foreground = '#000000'
    } else {
        $WPFisEncrypted_textBox.Background = '#FF6347'
        $WPFisEncrypted_textBox.Foreground = '#000000'
    }


    $lastSyncDateTimeLocalTime = Convert-UTCtoLocal (Get-Date $IntuneManagedDevice.lastSyncDateTime)
    #$lastSyncDateTimeLocalTime = Get-Date $lastSyncDateTimeLocalTime -Format 'yyyy-MM-dd HH:mm:ss'
    
    # This is used in textBox ToolTip
    $lastSyncDateTimeUFormatted = Get-Date $lastSyncDateTimeLocalTime -uformat '%G-%m-%d %H:%M:%S'
    $WPFlastSync_textBox.Tooltip = "Last Sync DateTime (yyyy-MM-dd HH:mm:ss): $lastSyncDateTimeUFormatted"

    $lastSyncDays = (New-Timespan $lastSyncDateTimeLocalTime).Days
    $lastSyncHours = (New-Timespan $lastSyncDateTimeLocalTime).Hours
   
    # This would be in UTC time with this syntax
    #$enrolledDays = (New-Timespan $IntuneManagedDevice.enrolledDateTime).Days

    if($lastSyncDays -le 30) {
        # Green
        $ForegroundColor           = '#000000'
        $BackgroundColor           = '#7FFF00'
    } elseif (($lastSyncDays -gt 30) -and ($lastSyncDays -le 90)) {
        # Yellow
        $ForegroundColor           = '#000000'
        $BackgroundColor           = 'yellow'
    } elseif (($lastSyncDays -gt 90) -and ($lastSyncDays -le 180)) {
        # Light red
        $ForegroundColor           = '#000000'
        $BackgroundColor           = '#FFB1B1'
    } elseif ($lastSyncDays -gt 180) {
        # Red
        $ForegroundColor           = '#000000'
        $BackgroundColor           = '#FF6347'
    }

    if($lastSyncDays -eq 0) {
        $WPFlastSync_textBox.Text = "$lastSyncHours hours ago"
    } else {
        $WPFlastSync_textBox.Text = "$lastSyncDays days ago"
    }

    $WPFlastSync_textBox.Foreground = $ForegroundColor
    $WPFlastSync_textBox.Background = $BackgroundColor

    ##### Windows Autopilot information #####
    $WPFAutopilotEnrolled_textBox.Text = $IntuneManagedDevice.autopilotEnrolled

    # Get Windows Autopilot GroupTag and assigned profile if device is autopilotEnrolled
    if($IntuneManagedDevice.autopilotEnrolled) {
        # Find device from Windows Autopilot
        $url = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,%27$($IntuneManagedDevice.serialNumber)%27)&_=1577625591868"
        $AutopilotDevice = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

        if($AutopilotDevice) {

            # Get more detailed information including Windows Autopilot IntendedProfile
            $url = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities/$($AutopilotDevice.id)?`$expand=deploymentProfile,intendedDeploymentProfile&_=1578315612557"
            $AutopilotDeviceWithAutpilotProfile = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

            if($AutopilotDeviceWithAutpilotProfile) {

                $WPFAutopilotGroupTag_textBox.Text = $AutopilotDeviceWithAutpilotProfile.groupTag
                
                # Check that assigned Autopilot Profile is same than Intended Autopilot Profile
                if($AutopilotDeviceWithAutpilotProfile.deploymentProfile.id -eq $AutopilotDeviceWithAutpilotProfile.intendedDeploymentProfile.id) {
                    $WPFAutopilotProfile_textBox.Text = $AutopilotDeviceWithAutpilotProfile.deploymentProfile.displayName
                } else {
                    # Windows Autopilot current and intended profile are different so we have to wait Autopilot sync to fix that
                    $WPFAutopilotProfile_textBox.Text = 'Profile sync is not ready'
                }

                $WPFAutopilotProfile_textBox.Tooltip = $AutopilotDeviceWithAutpilotProfile | ConvertTo-Json
                $WPFAutopilotGroupTag_textBox.Tooltip = $AutopilotDeviceWithAutpilotProfile | ConvertTo-Json
            }
        }
    }

    # Get Intune device primary user
    $url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($IntuneDeviceId)/users"
    $primaryUserId = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
    $primaryUserId = $primaryUserId.id

    # Set Primary User id to zero if we got nothing. This is used later with Application assignment report information
    if($primaryUserId -eq $null) {
        $primaryUserId = '00000000-0000-0000-0000-000000000000'
    }

    $UserGroupsMemberOf = $null

    if($primaryUserId -ne '00000000-0000-0000-0000-000000000000') {
    
        $user = $null

        # Check we have valid GUID
        if([System.Guid]::Parse($primaryUserId)) {
            # Get user information
            $url = "https://graph.microsoft.com/beta/users/$($primaryUserId)?`$select=id,displayName,mail"
            $user = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
        }

        if($user) {
            $WPFprimaryUser_textBox.Text = $user.mail

            $url = "https://graph.microsoft.com/beta/users/$($primaryUserId)/memberOf?_=1577625591876"

            $UserGroupsMemberOf = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
            if($UserGroupsMemberOf) {
    
                # Show user Group Membership on PrimaryUser Tooltip
                $PrimaryUserToolTip = "User is member of AzureAD Groups:`n"
                $UserGroupsMemberOf | Sort-Object displayName | Foreach-Object { $PrimaryUserToolTip += "* $($_.displayName)`n" }
                $WPFprimaryUser_textBox.Tooltip = $PrimaryUserToolTip

            } else {
                Write-Host "Did not find any groups for user $($primaryUserId)"
            }
        } else {
            $WPFprimaryUser_textBox.Text = 'user not found'
        }
    } else {
        $WPFprimaryUser_textBox.Text = 'Shared device'
    }


    # Get Logged on users information
    [String]$usersLoggedOn = @()
    foreach($LoggedOnUser in $IntuneManagedDevice.usersLoggedOn) {
        # Check we have valid GUID
        if([System.Guid]::Parse($LoggedOnUser.userId)) {
            # Get user information
            $url = "https://graph.microsoft.com/beta/users/$($LoggedOnUser.userId)?`$select=id,displayName,mail"
            $AADuser = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
        }
        $usersLoggedOn += "$($AADuser.mail)`n"
        $lastLogOnDateTime = Get-Date $LoggedOnUser.lastLogOnDateTime -Format 'yyyy-MM-dd HH:mm:ss'
        $usersLoggedOn += "$lastLogOnDateTime`n`n"
    }
    $WPFIntuneDeviceDetails_RecentCheckins_textBox.Text = $usersLoggedOn

    # Get Device AzureADGroup memberships
    $url = "https://graph.microsoft.com/v1.0/devices?`$filter=deviceid%20eq%20`'$($IntuneManagedDevice.azureADDeviceId)`'"
    $AzureADDevice = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

    if($AzureADDevice) {

        $url = "https://graph.microsoft.com/v1.0/devices/$($AzureADDevice.id)/memberOf?_=1577625591876"
        $deviceGroupMemberships = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

        if($deviceGroupMemberships) {
            [String]$GroupNames = $null
            foreach($group in $deviceGroupMemberships | Sort-Object displayName) {
                $GroupNames += "* $($group.displayName)`n"
            }
            $WPFIntuneDeviceDetails_GroupMemberships_textBox.Text = $GroupNames
        } else {
            Write-Output "Did not find any groups"
        }
    }

    # Device json information

    # Copy additional information attributes
    $IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name activationLockBypassCode -Value "$($AdditionalDeviceInformation.activationLockBypassCode)" -Force
    $IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name iccid -Value "$($AdditionalDeviceInformation.iccid)" -Force
    $IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name udid -Value "$($AdditionalDeviceInformation.udid)" -Force
    $IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name roleScopeTagIds -Value "$($AdditionalDeviceInformation.roleScopeTagIds)" -Force
    $IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name ethernetMacAddress -Value "$($AdditionalDeviceInformation.ethernetMacAddress)" -Force
    $IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name processorArchitecture -Value "$($AdditionalDeviceInformation.processorArchitecture)" -Force
    $IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name hardwareInformation -Value "$($AdditionalDeviceInformation.hardwareInformation)" -Force

    $WPFIntuneDeviceDetails_json_textBox.Text = $IntuneManagedDevice | ConvertTo-Json -Depth 4

    # Get Application Assignments
    $AppsWithAssignments = Get-ApplicationsWithAssignments

    # Get all applications targeted to specific user AND device
    # if there is no Primary User then we get only device targeted applications
    # We will get all device AND user targeted apps. We will need to figure out which apps came from which AzureAD Group targeting

	# Intune original request
    #$url = "https://graph.microsoft.com/beta/users('$($IntuneManagedDevice.userId)')/mobileAppIntentAndStates('$IntuneDeviceId')"

	# Using Primary User id
    $url = "https://graph.microsoft.com/beta/users('$($primaryUserId)')/mobileAppIntentAndStates('$IntuneDeviceId')"

    # Send MSGraph request
    $mobileAppIntentAndStates = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
    if($mobileAppIntentAndStates) {
        $mobileAppIntentAndStatesMobileAppList = $mobileAppIntentAndStates.mobileAppList
    } else {
        $mobileAppIntentAndStatesMobileAppList = $null
    }

    $script:AppsAssignmentsObservableCollection = @()

    # We duplicate $mobileAppIntentAndStatesMobileAppList and remove every app in every foreach loop if we find assignments
    # If for any reason we don't find Assignment targeted to either All Users, All Devices, member of Device group or member of Primary User group
    # then we will know that at the end
    # Case could be that we have nested groups targeting so we actually don't know which nested group was original assignment reason
    # So far we will report source to be unknown. Maybe in the future releases we can do recursive group membership search
    # to find out why application is targeted to device (this could actually be very important information)
    $CopyOfMobileAppIntentAndStatesMobileAppList = $mobileAppIntentAndStatesMobileAppList

    $odatatype = $null
    $assignmentGroup = $null 

    # Go through all Application Assignments for this specific device and primary user
    # Create new object (for ListView) when ever we find assignment targeted to device and/or user
    foreach ($mobileAppIntentAndStatesMobileApp in $mobileAppIntentAndStatesMobileAppList) {

        $assignmentGroup = 'unknown'
        $AppHadAssignments = $false
        $displayName = $null
        $properties = $null

        # Get Application information with Assignment details
        # Get it once and use it many times
        $App = $AppsWithAssignments | Where-Object { $_.id -eq "$($mobileAppIntentAndStatesMobileApp.applicationId)" }
        #$App

        # Remove #microsoft.graph. from @odata.type
        $odatatype = $App.'@odata.type'
        $odatatype = $odatatype.Replace('#microsoft.graph.', '')

        if ($App.licenseType -eq 'offline') {
            $displayName = "$($App.displayname) (offline)"
        }
        else {
            $displayName = "$($App.displayname)"
        }

        # Go through all Assignments in Application
        # Notice we can have at least 4 different Assignments showing here so we actually need to check every Assignment
        # All Users, All Devices, group specific included, group specific excluded
        # And Available and Required types of assignment
        # Excluded assignments are not available for All Users and All Devices
        Foreach ($Assignment in $App.Assignments) {
            
             # We will see Assignment which are not targeted to this device so we need to exclude those out
            $IncludeApplicationAssignmentInSummary = $false
            $context = '_unknown'
            $assignmentGroup = $null

            if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
                # Special case for All Users
                $assignmentGroup = 'All Users'
                $context = 'User'

                $IncludeApplicationAssignmentInSummary = $true
            }
            
            if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
                # Special case for All Devices
                $assignmentGroup = 'All Devices'
                $context = 'Device'

                $IncludeApplicationAssignmentInSummary = $true
            }
            
            if (($Assignment.target.'@odata.type' -ne '#microsoft.graph.allLicensedUsersAssignmentTarget') -and ($Assignment.target.'@odata.type' -ne '#microsoft.graph.allDevicesAssignmentTarget')) {

                # Group based assignment. We need to get AzureAD Group Name
                # #microsoft.graph.groupAssignmentTarget

                # Test if device is member of this group
                if($deviceGroupMemberships | Where-Object { $_.id -eq $Assignment.target.groupId}) {
                    $context = 'Device'
                    $assignmentGroup = $deviceGroupMemberships | Where-Object { $_.id -eq $Assignment.target.groupId} | Select-object -ExpandProperty displayName

                    $IncludeApplicationAssignmentInSummary = $true
                } else {
                    # Group not found on member of devicegroups
                }

                # Test if primary user is member of assignment group
                if($UserGroupsMemberOf | Where-Object { $_.id -eq $Assignment.target.groupId}) {
                    if($assignmentGroup) {
                        # Device also is member of this group. Now we got mixed User and Device memberships
                        # Maybe not good practise but it is possible

                        $context = '_Device/User'
                    } else {
                        $context = 'User'
                        $assignmentGroup = $UserGroupsMemberOf | Where-Object { $_.id -eq $Assignment.target.groupId} | Select-object -ExpandProperty displayName
                    }
                    $IncludeApplicationAssignmentInSummary = $true
                } else {
                    # Group not found on member of devicegroups
                }
            }

            if($IncludeApplicationAssignmentInSummary) {

                # Set included/excluded attribute
                $AppIncludeExclude = ''
                if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget') {
                    $AppIncludeExclude = 'Included'
                }
                if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget') {
                    $AppIncludeExclude = 'Excluded'
                }

                $assignmentIntent = $Assignment.intent
				
				if(($assignmentIntent -eq 'available') -and ($mobileAppIntentAndStatesMobileApp.installState -eq 'unknown')) {
					$mobileAppIntentAndStatesMobileApp.installState = 'Available for install'
				}

				if(($assignmentIntent -eq 'required') -and ($mobileAppIntentAndStatesMobileApp.installState -eq 'unknown')) {
					$mobileAppIntentAndStatesMobileApp.installState = 'Waiting for install status'
				}

				$assignmentFilterId = $Assignment.target.deviceAndAppManagementAssignmentFilterId
				$assignmentFilterDisplayName = $AllIntuneFilters | Where-Object { $_.id -eq $assignmentFilterId } | Select-Object -ExpandProperty displayName
				
				$FilterMode = $Assignment.target.deviceAndAppManagementAssignmentFilterType
				if($FilterMode -eq 'None') {
					$FilterMode = $null
				}


                $properties = @{
                    context                          = $context
                    odatatype                        = $odatatype
                    displayname                      = $displayName
                    version                          = $mobileAppIntentAndStatesMobileApp.displayVersion
                    assignmentIntent                 = $assignmentIntent
                    IncludeExclude                   = $AppIncludeExclude
                    assignmentGroup                  = $assignmentGroup
                    installState                     = $mobileAppIntentAndStatesMobileApp.installState
                    lastModifiedDateTime             = $App.lastModifiedDateTime
                    id                               = $App.id
					filter							 = $assignmentFilterDisplayName
					filterMode						 = $FilterMode
                }

                # Create new custom object every time inside foreach-loop
                # If you create custom object outside of foreach then you would edit same custom object on every foreach cycle resulting only 1 app in custom object array
                $CustomObject = New-Object -TypeName PSObject -Prop $properties

                # Add custom object to our custom object array.
                $script:AppsAssignmentsObservableCollection += $CustomObject
            }
        }
        
        # Remove App from our copy object array if any assignment was found
        $AppWithAssignment = $script:AppsAssignmentsObservableCollection | Where-Object { $_.id -eq $mobileAppIntentAndStatesMobileApp.applicationId }
        if ($AppWithAssignment) {
            # App had Assignment so we remove App from copy array

            # We will end up having only Apps which we did NOT find assignments in this array
            # This is reserved for possible future features
            $CopyOfMobileAppIntentAndStatesMobileAppList = $CopyOfMobileAppIntentAndStatesMobileAppList | Where-Object { $_.applicationId -ne $mobileAppIntentAndStatesMobileApp.applicationId}

        } else {
            # We could not determine Assignment source

            $context = '_unknown'
 
            # App Intent requiredInstall is different than App Assignment so we remove word Install
            $assignmentIntent = $mobileAppIntentAndStatesMobileApp.mobileAppIntent
            $assignmentIntent = $assignmentIntent.Replace('Install','')

            $AppIncludeExclude = ''
            $assignmentGroup = "unknown (possible nested group)"
    
            $properties = @{
                context                          = $context
                odatatype                        = $odatatype
                displayname                      = $displayName
                version                          = $mobileAppIntentAndStatesMobileApp.displayVersion
                assignmentIntent                 = $assignmentIntent
                IncludeExclude                   = $AppIncludeExclude
                assignmentGroup                  = $assignmentGroup
                installState                     = $mobileAppIntentAndStatesMobileApp.installState
                lastModifiedDateTime             = $App.lastModifiedDateTime
                id                               = $App.id
				filter							 = $null
				filterMode						 = $null
            }
            $CustomObject = New-Object -TypeName PSObject -Prop $properties
            $script:AppsAssignmentsObservableCollection += $CustomObject
        }
    }

    if($script:AppsAssignmentsObservableCollection.Count -gt 1) {
        # ItemsSource works if we are sorting 2 or more objects
        $WPFlistView_ApplicationAssignments.Itemssource = $script:AppsAssignmentsObservableCollection | Sort-Object -Property context, @{expression = 'assignmentIntent';descending = $true},IncludeExclude,displayName
    } else {
        # Only 1 object so we can't do sorting
        # If we try to sort here then our object array breaks and it does not work for ItemsSource
        $WPFlistView_ApplicationAssignments.Itemssource = $script:AppsAssignmentsObservableCollection
    }

##################################################################################################
    # Create Configuration Assignment information

    $script:ConfigurationsAssignmentsObservableCollection = @()

    # Get Device Configurations install state - OLD url
    #$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$IntuneDeviceId/deviceConfigurationStates"

    # Intune uses this Graph API url - OLD url
    #$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$IntuneDeviceId/deviceConfigurationStates?`$filter=((platformType%20eq%20'android')%20or%20(platformType%20eq%20'androidforwork')%20or%20(platformType%20eq%20'androidworkprofile')%20or%20(platformType%20eq%20'ios')%20or%20(platformType%20eq%20'macos')%20or%20(platformType%20eq%20'WindowsPhone81')%20or%20(platformType%20eq%20'Windows81AndLater')%20or%20(platformType%20eq%20'Windows10AndLater')%20or%20(platformType%20eq%20'all'))&_=1578056411936"
	
	# Intune uses this nowdays 20220330
	# Note this is post type request
	$url = "https://graph.microsoft.com/beta/deviceManagement/reports/getConfigurationPoliciesReportForDevice"

	$GraphAPIPostRequest = @"
{
    "select": [
        "IntuneDeviceId",
        "PolicyBaseTypeName",
        "PolicyId",
        "PolicyStatus",
        "UPN",
        "UserId",
        "PspdpuLastModifiedTimeUtc",
        "PolicyName",
        "UnifiedPolicyType"
    ],
    "filter": "((PolicyBaseTypeName eq 'Microsoft.Management.Services.Api.DeviceConfiguration') or (PolicyBaseTypeName eq 'DeviceManagementConfigurationPolicy') or (PolicyBaseTypeName eq 'DeviceConfigurationAdmxPolicy') or (PolicyBaseTypeName eq 'Microsoft.Management.Services.Api.DeviceManagementIntent')) and (IntuneDeviceId eq '$IntuneDeviceId')",
    "skip": 0,
    "top": 50
}
"@

	Write-Host "Get Intune device Configuration Assignment information"
	$MSGraphRequest = Invoke-MSGraphRequest -Url $url -Content $GraphAPIPostRequest.ToString() -HttpMethod 'POST'
	$Success = $?

	if($Success) {
		#Write-Host "Success"
	} else {
		# Invoke-MSGraphRequest failed
		Write-Error "Error getting Intune device Configuration Assignment information"
		return 1
	}

	# Get AllMSGraph pages
	# This is also workaround to get objects without assigning them from .Value attribute
	$ConfigurationPoliciesReportForDevice = Get-MSGraphAllPages -SearchResult $MSGraphRequest
	$Success = $?

	if($Success) {
		Write-Host "Success"
	} else {
		# Invoke-MSGraphRequest failed
		Write-Error "Error getting Intune device Configuration Assignment information"
		return 1
	}

	# Original
    #$deviceConfigurationStates = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

	$ConfigurationPoliciesReportForDevice = Objectify_JSON_Schema_and_Data_To_PowershellObjects $ConfigurationPoliciesReportForDevice
	
	# Sort policies by PolicyId so we will download policies only once in next steps
	$ConfigurationPoliciesReportForDevice = $ConfigurationPoliciesReportForDevice | Sort-Object -Property PolicyId
	
	# DEBUG to clipboard -> Paste to text editor after script has run
	#$ConfigurationPoliciesReportForDevice | ConvertTo-Json -Depth 6 | Set-Clipboard

    $lastDeviceConfigurationId = $null

    $CopyOfConfigurationPoliciesReportForDevice = $ConfigurationPoliciesReportForDevice
    $odatatype = $null
    $assignmentGroup = $null 

    foreach($ConfigurationPolicyReportState in $ConfigurationPoliciesReportForDevice) {

        $assignmentGroup = $null
        $DeviceConfiguration = $null
        $IntuneDeviceConfigurationPolicyAssignments = $null
        $IncludeConfigurationAssignmentInSummary = $true
        $properties = $null
		$odatatype = $ConfigurationPolicyReportState.UnifiedPolicyType_loc

		# Change PolicyStatus numbers to text
		Switch ($ConfigurationPolicyReportState.PolicyStatus) {
			1 { $ConfigurationPolicyReportState.PolicyStatus = 'Not applicable' }
			2 { $ConfigurationPolicyReportState.PolicyStatus = 'Succeeded' }   # User based result?
			3 { $ConfigurationPolicyReportState.PolicyStatus = 'Succeeded' }   # Device based result?
			5 { $ConfigurationPolicyReportState.PolicyStatus = 'Error' }   	   # User based result?
			6 { $ConfigurationPolicyReportState.PolicyStatus = 'Conflict' }
			Default { }
		}


		if($ConfigurationPolicyReportState.PolicyBaseTypeName -eq 'Microsoft.Management.Services.Api.DeviceManagementIntent') {
			# Endpoint Security templates information does not include assignments
			# So we get assignment information separately to those templates
			#https://graph.microsoft.com/beta/deviceManagement/intents/932d590f-b340-4a7c-b199-048fb98f09b2/assignments

			$url = "https://graph.microsoft.com/beta/deviceManagement/intents/$($ConfigurationPolicyReportState.PolicyId)/assignments"
			$IntuneDeviceConfigurationPolicyAssignments = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
		} else {
			$IntuneDeviceConfigurationPolicyAssignments = $IntuneConfigurationProfilesWithAssignments | Where-Object id -eq $ConfigurationPolicyReportState.PolicyId | Select-Object -ExpandProperty assignments
		}

		if($ConfigurationPolicyReportState.PolicyStatus -eq 'Not applicable' ) {
			$context = ''
		} else {
			# Default value started with.
			# This will change later on the script if we find where assignment came from
			$context = '_unknown'
		}

        $lastModifiedDateTime = $DeviceConfiguration.PspdpuLastModifiedTimeUtc

        # Remove #microsoft.graph. from @odata.type
        # "@odata.type": "#microsoft.graph.windows10GeneralConfiguration",
        #$odatatype = $DeviceConfiguration.'@odata.type'
        $odatatype = $odatatype.Replace('#microsoft.graph.', '')
        $assignmentGroup = $null

        foreach ($IntuneDeviceConfigurationPolicyAssignment in $IntuneDeviceConfigurationPolicyAssignments) {

            $assignmentGroup = $null

            # Only include Configuration which have assignments targeted to this device/user
            $IncludeConfigurationAssignmentInSummary = $false

            $context = '_unknown'
            
            if ($IntuneDeviceConfigurationPolicyAssignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
                # Special case for All Users
                $assignmentGroup = 'All Users'
                $context = 'User'

                $IncludeConfigurationAssignmentInSummary = $true
            }

            if ($IntuneDeviceConfigurationPolicyAssignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
                # Special case for All Devices
                $assignmentGroup = 'All Devices'
                $context = 'Device'

                $IncludeConfigurationAssignmentInSummary = $true
            }

            if (($IntuneDeviceConfigurationPolicyAssignment.target.'@odata.type' -ne '#microsoft.graph.allLicensedUsersAssignmentTarget') -and ($IntuneDeviceConfigurationPolicyAssignment.target.'@odata.type' -ne '#microsoft.graph.allDevicesAssignmentTarget')) {

                # Group based assignment. We need to get AzureAD Group Name
                # #microsoft.graph.groupAssignmentTarget

                # Test if device is member of this group
                if($deviceGroupMemberships | Where-Object { $_.id -eq $IntuneDeviceConfigurationPolicyAssignment.target.groupId}) {
                    $assignmentGroup = $deviceGroupMemberships | Where-Object { $_.id -eq $IntuneDeviceConfigurationPolicyAssignment.target.groupId} | Select-Object -ExpandProperty displayName
                    #Write-Host "device group found: $($assignmentGroup.displayName)"
                    $context = 'Device'

                    $IncludeConfigurationAssignmentInSummary = $true
                } else {
                    # Group not found on member of devicegroups
                }

                # Test if primary user is member of assignment group
                if($UserGroupsMemberOf | Where-Object { $_.id -eq $IntuneDeviceConfigurationPolicyAssignment.target.groupId}) {
                    if($assignmentGroup) {
                        # Device also is member of this group. Now we got mixed User and Device memberships
                        # Maybe not good practise but it is possible

                        $context = '_Device/User'
                    } else {
                        $assignmentGroup = $UserGroupsMemberOf | Where-Object { $_.id -eq $IntuneDeviceConfigurationPolicyAssignment.target.groupId} | Select-Object -ExpandProperty displayName
                        #Write-Host "User group found: $($assignmentGroup.displayName)"
                        $context = 'User'
                    }
                    $IncludeConfigurationAssignmentInSummary = $true
                } else {
                    # Group not found on member of devicegroups
                }
            }

            if($IncludeConfigurationAssignmentInSummary) {
            
                # Set included/excluded attribute
                $AppIncludeExclude = ''
                if ($IntuneDeviceConfigurationPolicyAssignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget') {
                    $AppIncludeExclude = 'Included'
                }
                if ($IntuneDeviceConfigurationPolicyAssignment.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget') {
                    $AppIncludeExclude = 'Excluded'
                }

				$state = $ConfigurationPolicyReportState.PolicyStatus

				$assignmentFilterId = $IntuneDeviceConfigurationPolicyAssignment.target.deviceAndAppManagementAssignmentFilterId
				$assignmentFilterDisplayName = $AllIntuneFilters | Where-Object { $_.id -eq $assignmentFilterId } | Select-Object -ExpandProperty displayName
				
				$FilterMode = $IntuneDeviceConfigurationPolicyAssignment.target.deviceAndAppManagementAssignmentFilterType
				if($FilterMode -eq 'None') {
					$FilterMode = $null
				}

                $properties = @{
                    context                          = $context
                    odatatype                        = $odatatype
                    userPrincipalName                = $ConfigurationPolicyReportState.UPN
                    displayname                      = $ConfigurationPolicyReportState.PolicyName
                    assignmentIntent                 = $assignmentIntent
                    IncludeExclude                   = $AppIncludeExclude
                    assignmentGroup                  = $assignmentGroup
                    state                            = $state
                    id                               = $ConfigurationPolicyReportState.PolicyId
					filter							 = $assignmentFilterDisplayName
					filterMode						 = $FilterMode
                }

                # Create new custom object every time inside foreach-loop
                # If you create custom object outside of foreach then you would edit same custom object on every foreach cycle resulting only 1 app in custom object array
                $CustomObject = New-Object -TypeName PSObject -Prop $properties

                # Add custom object to our custom object array.
                $script:ConfigurationsAssignmentsObservableCollection += $CustomObject
            }
        }

        # Remove DeviceConfiguration from our copy object array if any assignment was found
        $DeviceConfigurationWithAssignment = $script:ConfigurationsAssignmentsObservableCollection | Where-Object { $_.id -eq $ConfigurationPolicyReportState.PolicyId }
        if ($DeviceConfigurationWithAssignment) {
            # Remove DeviceConfiguration from copy array because that Configration had Assignment
            # We will end up only having Configurations which we did NOT find assignments
            # We may use this object array with future features
            $CopyOfConfigurationPoliciesReportForDevice = $CopyOfConfigurationPoliciesReportForDevice | Where-Object { $_.id -ne $ConfigurationPolicyReportState.PolicyId}

        } else {
            # We could not determine Assignment source

            $context = '_unknown'
            $AppIncludeExclude = ''
            #$assignmentGroup = "unknown (possible nested group)"
			$assignmentGroup = "unknown"
    
            $properties = @{
                context                          = $context
                odatatype                        = $odatatype
                userPrincipalName                = $ConfigurationPolicyReportState.UPN
                displayname                      = $ConfigurationPolicyReportState.PolicyName
                assignmentIntent                 = $assignmentIntent
                IncludeExclude                   = $AppIncludeExclude
                assignmentGroup                  = $assignmentGroup
                state                            = $ConfigurationPolicyReportState.PolicyStatus
                id                               = $ConfigurationPolicyReportState.PolicyId
				filter							 = $null
				filterMode						 = $null
            }

            $CustomObject = New-Object -TypeName PSObject -Prop $properties
            $script:ConfigurationsAssignmentsObservableCollection += $CustomObject
        }

        $lastDeviceConfigurationId = $ConfigurationPolicyReportState.PolicyId
    }
    
    if($script:ConfigurationsAssignmentsObservableCollection.Count -gt 1) {
        # ItemsSource works if we are sorting 2 or more objects
        $WPFlistView_ConfigurationsAssignments.Itemssource = $script:ConfigurationsAssignmentsObservableCollection | Sort-Object displayName,userPrincipalName
    } else {
        # Only 1 object so we can't do sorting
        # If we try to sort here then our object array breaks and it does not work for ItemsSource
        $WPFlistView_ConfigurationsAssignments.Itemssource = $script:ConfigurationsAssignmentsObservableCollection
    }

    # Status textBox
    $DateTime = Get-Date -Format 'yyyy-MM-dd HH:mm.ss'
    $WPFbottom_textBox.Text = "Device details updated $DateTime"
}


##################################################################################################

#region Form specific functions closing etc...

$WPFRefresh_button.Add_Click({
    Get-DeviceInformation
})


# Add Exit

$Form.Add_Closing{

    if ($ProductionMode) {

        #ONLY USE THIS IN PRODUCTION.
        # This works, Powershell process is (forcefully) killed when WPF ends
        [System.Windows.Forms.Application]::Exit()
        Stop-Process $pid

    }
}

#endregion Form specific functions closing etc...

##################################################################################################

#region Main

#===========================================================================
# Shows the form
#===========================================================================

# Initiate variables

[array]$script:AppsAssignmentsObservableCollection = @()

##### Application ListView sorting

$script:ApplicationGridLastColumnClicked = $null

# Add Application ListView Grid sorting code when column is pressed
[Windows.RoutedEventHandler]$ApplicationsGridColumnClickEventCode = {

    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($WPFlistView_ApplicationAssignments.Itemssource)

    # Not used because we will be confused if some column was sorted earlier
    # because this remembers that column last sort ordering
    #$sort = $view.SortDescriptions[0].Direction

    $columnHeader = $_.OriginalSource.Column.Header

    if ($columnHeader -eq $null) { return }

    if ($columnHeader -eq $script:ApplicationGridLastColumnClicked) {
        # Same column clicked so reversing sorting order
        $sort = $view.SortDescriptions[0].Direction
        $direction = if ($sort -and 'Descending' -eq $sort) { 'Ascending' } else { 'Descending' }
    } else {
        # New column clicked so we always sort Ascending first time column is clicked

        # Always do Ascending sort unless we have pressed same column twice or more times
        $direction = 'Ascending'
    }

    #Write-Verbose "$($_.OriginalSource.Column.Header) header clicked, doing $direction sorting to column/table."

    $view.SortDescriptions.Clear()
    $sortDescription = New-Object System.ComponentModel.SortDescription($columnHeader, $direction)
    $view.SortDescriptions.Add($sortDescription)

    # Save info we clicked this column.
    # If we click next time same column then we just reverse sort order
    $script:ApplicationGridLastColumnClicked = $columnHeader
}

# Add Application ListView Grid column select event handler for sorting when grid column is clicked
$WPFlistView_ApplicationAssignments.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $ApplicationsGridColumnClickEventCode)

#####

[array]$script:ConfigurationsAssignmentsObservableCollection = @()

##### Configurations ListView sorting

$script:ConfigurationProfilesGridLastColumnClicked = $null

# Add ConfigurationProfiles ListView Grid sorting code when column is pressed
[Windows.RoutedEventHandler]$ConfigurationProfilesGridColumnClickEventCode = {

    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($WPFlistView_ConfigurationsAssignments.Itemssource)

    # Not used because we will be confused if some column was sorted earlier
    # because this remembers that column last sort ordering
    #$sort = $view.SortDescriptions[0].Direction

    $columnHeader = $_.OriginalSource.Column.Header

    if ($columnHeader -eq $null) { return }

    if ($columnHeader -eq $script:ConfigurationProfilesGridLastColumnClicked) {
        # Same column clicked so reversing sorting order
        $sort = $view.SortDescriptions[0].Direction
        $direction = if ($sort -and 'Descending' -eq $sort) { 'Ascending' } else { 'Descending' }
    } else {
        # New column clicked so we always sort Ascending first time column is clicked

        # Always do Ascending sort unless we have pressed same column twice or more times
        $direction = 'Ascending'
    }

    #Write-Verbose "$($_.OriginalSource.Column.Header) header clicked, doing $direction sorting to column/table."

    $view.SortDescriptions.Clear()
    $sortDescription = New-Object System.ComponentModel.SortDescription($columnHeader, $direction)
    $view.SortDescriptions.Add($sortDescription)

    # Save info we clicked this column.
    # If we click next time same column then we just reverse sort order
    $script:ConfigurationProfilesGridLastColumnClicked = $columnHeader
}

# Add ConfigurationProfiles ListView Grid column select event handler for sorting when grid column is clicked
$WPFlistView_ConfigurationsAssignments.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $ConfigurationProfilesGridColumnClickEventCode)

#####


$TenantId = $null

# Howto convert images to base64 and back in XAML
# http://vcloud-lab.com/entries/powershell/powershell-gui-encode-decode-images

$logo = @'
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EE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAADhDwEGADG6OdrIk7dVAAAAAElFTkSuQmCC
'@
$WPFbottomRightLogoimage.source = DecodeBase64Image -ImageBase64 $Logo

# Compliant.png
$DeviceCompliantPictureBase64 = @'
iVBORw0KGgoAAAANSUhEUgAAAA8AAAAPCAIAAAC0tAIdAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAEnQAABJ0Ad5mH3gAAAGMSURBVChTlZDPSwJBHMXn1iW6dojoVhbUKbJTf1DZJVIz06AoCpWirSil6FgdUsjfkmju6prkxS4ZmdFKUh0kWGp1xd66g+kpegxfdr7v8515s6TxH3XQtbqUKGw4eS3DalzpSfbJJtdr1Gvql84KJ+Yg0XvJYoCYA0qd95KlILkrn1OiRWdLRzo3sYSINUSWw3Rhi7EZN8mVz1RMoSVZNAUUr8WpKKot1qtcFSRyvUrp6IMVAaxNQj0bddZDhMrN53fJ6FciJQqblN5LjeIAoCuRLnwAhR3OG2CtXXUjOppOfoLSW9cDyIcwovSWez2dviD7yTH0jzNTRp9yFZLssIOU3uVGMI3FsEPYckU7arLomLukwWAd8uOUjuQNep9iIKI93oeOUOFbKCpeFXtcpbRY/cCN6k/AgCutXY/2IC62WOgjZFUWKQ2lnrfxE9QBvMHShuo85FZwqRilIQzgYINPofEsVMRb8JPMywEl2mlIlN5D93qG0zji/Qw3HMmbvmoV6jXVQf+hRuMHSI7TQY1Q5WEAAAAASUVORK5CYII=
'@
$CompliantPictureObject = DecodeBase64Image -ImageBase64 $DeviceCompliantPictureBase64

# NotCompliant.png
$DeviceNotCompliantPictureBase64 = @'
iVBORw0KGgoAAAANSUhEUgAAAA8AAAAPCAIAAAC0tAIdAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAEnQAABJ0Ad5mH3gAAAFiSURBVChTY/hPCsCi+uepsz927f/z6AmUjwRQVL9Pzn/KwAtEzxgEnjJwvBDX+L5pB1QODBCqX6qZP2MQes6v+JxD+hmT6HM+hefcck8ZOD/3z4CqgKt+4x7yjFHkhaAyUOm7mIwvk2e/VDAEaeBXfMrA9fPYKYgykOo/L14+ZWB/LqAEVP2MWexzWz9Q8IWYxnNeBYj+19aeYMVg1V8mzX7GIAyUAKlmEf+YVwUUfMYhDTQYpFpACegesGKw6k+VzUAjIaqfs0m+T8gBCj5l4IfYBkRA1X9//YKqRjb7OZfsa2OnH1t3PWXgg4qgmY3iblaJj0W1f99/fCGhCfQlSDXQ3VZI7gaCN27QMHnGKv6ppg0o8lJWD+hLaJgcRQoTCHipagoNbzZJUHjzykPDu286VAWyaiB4n5T7lIEHHJ38oLgUU/++cTtUDgxQVEPAz5PgdPLwMZSPBLCoxgn+/wcA/sTjRbfvP80AAAAASUVORK5CYII=
'@
$NotCompliantPictureObject = DecodeBase64Image -ImageBase64 $DeviceNotCompliantPictureBase64

$Icon = @'
iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAACXBIWXMAAAsTAAALEwEAmpwYAAAAB3RJTUUH4gsYFyALbhoWlwAAAB1pVFh0Q29tbWVudAAAAAAAQ3JlYXRlZCB3aXRoIEdJTVBkLmUHAAACCklEQVRYw+1XQY7aMBR9tgMJ6mjaEaNEgnYUFVQOwCGQKqaaG3AfJPZcgAWrLGCLOAArtsABYAkSILCdPwuaEBRIS7NAo/JWcfL9/fz9//MPIyLCDcFxY9wJGAAwn88xmUwgxJGP7xMqlQps2061wGKxwGw2A3BMNcYYCoUiXNcFtNbU6XSoWCxQuVyicrlEpdJ3enn5Rp7nUVp0u93Qb+C7UvlBzWaTiIgMAFBKYb1eIygIIoKUElqr1CFWSmGz2YRjIoJhGJBSHo+AMQYhBDjnodHhmaUmEPiObo5zDsbYB6gC08ymXiCbzf65Ci6FzfM8TKdTKPVvuSCEwHg8RpLYGkmTB4MBhsNhqgj4vn99BAJoraG1/g+UMAnnzi8ooTS2f0XANE1kMpnY+6hoRXPGNM1QSwJIKbHb7a4noJTG29tPVKtVKCVPdtRqtbBarcLdaa3hui5eX3/h+TkfJp4QAqPRCL1e7zoCRATf16jVaqjX67Hv7XYby+XyhECx+BWNRgOO45zYPj5+Rr/fv1iKiUl4KXREFDtb39fYbrcx2/1+f29I7gTuBD4IgWt/kKL2RISk6Yn9AGMMuVwOjuPAsqwTDY+Oo8jn89Bah2IkpcTT0xcIIWK2lmXBtu0TEkIIPDx8Okj77+451jgQAULws7fZ+R6BgXMWsz/nO7hTOOcHAvckvCXeAQGrOg1kdasJAAAAAElFTkSuQmCC
'@
# Set Windows icon (upper left)
$Form.Icon = DecodeBase64Image -ImageBase64 $Icon

# Set taskbar icon
$Form.TaskbarItemInfo.Overlay = DecodeBase64Image -ImageBase64 $Icon
$Form.TaskbarItemInfo.Description = $Form.Title


#####

# Export images to cache folder
# This is workaround until I figure out how to use image objects in Devices ListView GridView cell

<#
# Try-catch because files are usually locked when using script again in same session
try {
    $return = Convert-Base64ToFile $DeviceCompliantPictureBase64 "$PSScriptRoot\cache\Compliant.png"
    Write-Verbose "Convert-Base64ToFile $PSScriptRoot\cache\Compliant.png $return"

    $return = Convert-Base64ToFile $DeviceNotCompliantPictureBase64 "$PSScriptRoot\cache\NotCompliant.png"
    Write-Verbose "Convert-Base64ToFile $PSScriptRoot\cache\NotCompliant.png $return"
} catch {}
#>

#####

# Gather information to show on UI
Get-DeviceInformation

#####

# In ProductionMode Powershell console Window is hidden and Powershell process is killed on exit
#$ProductionMode = $true
$ProductionMode = $false

if ($ProductionMode) {
    # Make PowerShell Disappear 
    $windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' 
    $asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru 
    $null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0) 
}

# This should be more responsive approach. We have not configured to run GUI in its own thread
# https://blog.netnerds.net/2016/01/showdialog-sucks-use-applicationcontexts-instead/
#$app = [Windows.Application]::new()
#$app.Run($Form)

# Show Form
$Form.ShowDialog() | out-null

