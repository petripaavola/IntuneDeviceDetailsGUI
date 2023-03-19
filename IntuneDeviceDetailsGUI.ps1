<#
.Synopsis
   Intune Device Details GUI ver 2.974
   

   Author:
   Petri.Paavola@yodamiitti.fi
   Modern Management Principal
   Microsoft MVP - Windows and Devices for IT
   
   2023-03-19
   
   https://github.com/petripaavola/IntuneDeviceDetailsGUI
.DESCRIPTION
   This tool visualizes Intune device and user details and
   Applications and Configurations Deployments.

   Tool includes search capability so you can just run script and find devices from
   script's built-in search.

   Some of this information is not shown easily or at all in Intune web console
   - Application and Configuration Deployments include information of
     what Azure AD Group was used for assignment and if filter was applied
   - Number of affected devices and/or users is shown with deployments
   - Last signed in users are shown
   - With shared devices you can select logged in user to show Application Deployments
   - JSON data inside Intune helps for example to build Azure AD Dynamic groups rules
   - This tool helps to understand why some Apps or Configuration Profiles are applying to device
     (what Azure AD group and/or filter is applied)

   Tip: hover with mouse on top of different values to get ToolTips.
        There is more info shown (for example DeviceName, PrimaryUser, OS/version)
	
   Test right click menus -> You can open MEM/Intune web console for many resources:
   Intune Device, Azure AD Device, PrimaryUser, Azure AD Groups, Applications, Filters, etc...
.EXAMPLE
   .\IntuneDeviceDetailsGUI.ps1
.EXAMPLE
   .\IntuneDeviceDetailsGUI.ps1 -Verbose
.EXAMPLE
    .\IntuneDeviceDetailsGUI.ps1 -Id 2e6e1d5f-b18a-44c6-989e-9bbb1efafbff
.EXAMPLE
    .\IntuneDeviceDetailsGUI.ps1 -IntuneDeviceId 2e6e1d5f-b18a-44c6-989e-9bbb1efafbff
.EXAMPLE
	Get-IntuneManagedDevice -Filter "devicename eq 'MyLoveMostPC'" | .\IntuneDeviceDetailsGUI.ps1
	
	Pipe Intune objects to script from Powershell console
.EXAMPLE
    Get-IntuneManagedDevice | Out-GridView -OutputMode Single | .\IntuneDeviceDetailsGUI.ps1
	
	Show Intune Devices in Out-GridView and pipe selected device to GUI
.INPUTS
   Intune DeviceId as a string or Intune Device object (which has id property)
.OUTPUTS
   None
.NOTES
.LINK
   https://github.com/petripaavola/IntuneDeviceDetailsGUI
#>


[CmdletBinding(DefaultParameterSetName = 'id')]
Param(
    [Parameter(Mandatory=$false,
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
    [String]$id = $null
)

$ScriptVersion = "ver 2.974"
$IntuneDeviceId = $id
$TimeOutBetweenGraphAPIRequests = 300

# Workaround to make variable scope to script
# these variables are used with events to get access to needed objects
# Future real fix is to use SyncHash variable with threading
$Script:IntuneManagedDevice = $null
$Script:AutopilotDeviceWithAutpilotProfile = $null
$Script:PrimaryUser = $null
$Script:AzureADDevice = $null
$script:ComboItemsInSearchComboBoxSource = $null
$script:SelectedIntuneDeviceForReportTextboxSource = $null
$Script:PrimaryUserGroupsMemberOf = $null
$Script:LatestCheckedInUserGroupsMemberOf = $null

# Reload cache automatically every Nth days
$Script:ReloadCacheEveryNDays = 1

# Limit Graph API Search results to top X devices specified here
$GraphAPITop = 100


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
        Title="IntuneDeviceDetails $ScriptVersion" Height="1000" MinHeight="800" Width="1200" MinWidth="1127" WindowStyle="ThreeDBorderWindow">
    <Window.TaskbarItemInfo>
        <TaskbarItemInfo/>
    </Window.TaskbarItemInfo>
    <Grid x:Name="IntuneDeviceDetails" Background="#FFE5EEFF">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="200" MinWidth="100" MaxWidth="400"/>
			<ColumnDefinition MinWidth="3" MaxWidth="3"/>
            <ColumnDefinition Width="*" MinWidth="100"/>
			<ColumnDefinition MinWidth="3" MaxWidth="3"/>
            <ColumnDefinition Width="550" MinWidth="200" MaxWidth="800"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="*" MinHeight="100" MaxHeight="100"/>
            <RowDefinition MinHeight="105" MaxHeight="105"/>
            <RowDefinition MinHeight="150"/>
			<RowDefinition MinHeight="3" MaxHeight="3"/>
            <RowDefinition MinHeight="150"/>
			<RowDefinition MinHeight="3" MaxHeight="3"/>
            <RowDefinition MinHeight="150"/>
            <RowDefinition Height="50" MinHeight="50" MaxHeight="50"/>
        </Grid.RowDefinitions>
        <Border x:Name="IntuneDeviceDetailsBorderTop" Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="5" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid x:Name="GridIntuneDeviceDetailsBorderTop" Grid.Row="0" ShowGridLines="False">
                <!-- <Button x:Name="Refresh_button" Content="Refresh" HorizontalAlignment="Right" VerticalAlignment="Bottom" Width="75" Margin="0,0,10,0"/> -->
				<Label x:Name="label_ConnectedAsUser" Content="Connected as:" HorizontalAlignment="Left" Margin="6,5,0,0" VerticalAlignment="Top" Height="28" FontWeight="Bold"/>
				<Label x:Name="label_ConnectedAsUser_UserName" Content="" HorizontalAlignment="Left" Margin="88,4,0,0" VerticalAlignment="Top" Height="28" Width="300" FontWeight="Bold" FontSize="14" Foreground="#FF004CFF"/>
                <Label x:Name="IntuneDeviceDetails_Label_DeviceName" Content="Showing information from device" HorizontalAlignment="Left" Margin="6,0,0,30" VerticalAlignment="Bottom" Height="26" FontWeight="Bold"/>
				<TextBox x:Name="IntuneDeviceDetails_textBox_DeviceName" HorizontalAlignment="Left" Height="27" Margin="10,0,0,5" Text="DeviceName" VerticalAlignment="Bottom" Width="380" FontWeight="Bold" FontSize="20" Foreground="#FF004CFF" IsReadOnly="True" ToolTipService.ShowDuration="60000">
					<TextBox.ContextMenu>
						<ContextMenu>
							<MenuItem x:Name='IntuneDeviceDetails_textBox_DeviceName_Menu_Copy' Header='Copy'/>
							<MenuItem x:Name='IntuneDeviceDetails_textBox_DeviceName_Menu_OpenDeviceInBrowser' Header='Open Intune Device in browser'/>
							<MenuItem x:Name='IntuneDeviceDetails_textBox_DeviceName_Menu_OpenAzureADDeviceInBrowser' Header='Open Azure AD Device in browser'/>
							<MenuItem x:Name='IntuneDeviceDetails_textBox_DeviceName_Menu_OpenAutopilotDeviceInBrowser' Header='Open Autopilot devices in browser and paste device serial to search'/>
						</ContextMenu>
					</TextBox.ContextMenu>
					<TextBox.ToolTip>
						<StackPanel>
							<TextBlock x:Name='IntuneDeviceDetails_textBox_DeviceName_ToolTip_DeviceName' FontWeight="Bold" FontSize="14" Margin="0,0,0,5"></TextBlock>
							<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,2" />
							<TextBlock FontSize="14" x:Name='IntuneDeviceDetails_textBox_DeviceName_ToolTip_DeviceProperties' FontFamily="Consolas"></TextBlock>
							<TextBlock/>
							<TextBlock FontSize="14" FontWeight="Bold">AzureAD device extensionAttributes</TextBlock>
							<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
							<TextBlock FontSize="14" x:Name='IntuneDeviceDetails_textBox_DeviceName_ToolTip_extensionAttributes' FontFamily="Consolas"></TextBlock>
						</StackPanel>
					</TextBox.ToolTip>
				</TextBox>
				<Label x:Name="label_GridIntuneDeviceDetailsBorderTop_Search" Content="Type search text or select Quick Filter from dropdown" HorizontalAlignment="Left" Margin="400,0,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
				<ComboBox x:Name="comboBox_GridIntuneDeviceDetailsBorderTop_Search" HorizontalAlignment="Left" Margin="400,24,0,0" VerticalAlignment="Top" Width="480" IsEditable="True" FontFamily="Consolas" FontSize="14">
				</ComboBox>
				<Image x:Name="GridIntuneDeviceDetailsBorderTop_image_Search_X" HorizontalAlignment="Left" Height="17" Margin="880,26,0,0" VerticalAlignment="Top" Width="17"/>
				<Button x:Name="Button_GridIntuneDeviceDetailsBorderTop_Search" Content="1. Search devices" HorizontalAlignment="Left" VerticalAlignment="Top" Width="100" Margin="910,25,0,0"/>
				<Label x:Name="label_GridIntuneDeviceDetailsBorderTop_CreateReport" Content="Selected device for report" HorizontalAlignment="Left" Margin="400,0,0,25" VerticalAlignment="Bottom" Height="26" FontWeight="Bold"/>
				<Label x:Name="label_GridIntuneDeviceDetailsBorderTop_FoundXDevices" Content="Found X devices" HorizontalAlignment="Left" Margin="755,0,0,24" VerticalAlignment="Bottom" Height="26" Width="120" Visibility="Hidden" Foreground="#FF004CFF" HorizontalContentAlignment="Right"/>
				<ComboBox x:Name="ComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport" HorizontalAlignment="Left" Height="23" Margin="400,0,0,5" Text="{Binding 'deviceName'}" VerticalAlignment="Bottom" Width="480" IsReadOnly="True" IsEditable="False" FontFamily="Consolas" FontSize="14" Focusable="False" ToolTipService.ShowDuration="60000">
					<ComboBox.ItemContainerStyle>
						<Style>
							<Setter Property="Control.ToolTip">
								<Setter.Value>  
									<StackPanel>  
										<TextBlock FontSize="14" FontFamily="Consolas" Text="{Binding SearchResultToolTip}" ToolTipService.ShowDuration="60000"/>
									</StackPanel>  
								</Setter.Value>  
							</Setter>  
						</Style>
					</ComboBox.ItemContainerStyle>
				</ComboBox>
				<Image x:Name="GridIntuneDeviceDetailsBorderTop_image_CreateReport_X" HorizontalAlignment="Left" Height="17" Margin="880,0,0,7" VerticalAlignment="Bottom" Width="17"/>
				<Button x:Name="Button_GridIntuneDeviceDetailsBorderTop_CreateReport" Content="2. Create report" HorizontalAlignment="Left" VerticalAlignment="Bottom" Width="100" Margin="910,0,0,7" IsEnabled="False" />
				<CheckBox x:Name="checkBox_ReloadCache" Content="Reload cache" HorizontalAlignment="Left" Margin="1015,0,0,19" VerticalAlignment="Bottom" Width="94" IsChecked="False" ToolTipService.ShowDuration="60000"/>
				<CheckBox x:Name="checkBox_SkipAppAndConfigurationAssignmentReport" Content="Basic info only (Skip Apps and Configurations Assignments)" HorizontalAlignment="Left" Margin="1015,0,0,3" VerticalAlignment="Bottom" Width="350" IsChecked="False" ToolTipService.ShowDuration="60000"/>
				<Label x:Name="label_UnknownAssignmentGroupsFoundWarningText" HorizontalAlignment="Left" Margin="1030,0,0,29" VerticalAlignment="Bottom"  FontWeight="Bold" Foreground="Red" Width="280" Visibility="Collapsed">
					<TextBlock TextWrapping="Wrap" Text="Unknown assignments found. Try enabling Reload cache and run report again"/>
				</Label>
            </Grid>
        </Border>
        <Border x:Name="IntuneDeviceDetailsBorderDeviceDetails" Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="5" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
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
                <TextBox x:Name="Compliance_textBox" HorizontalAlignment="Left" Height="23" Margin="600,4,0,0" Text="Compliance" VerticalAlignment="Top" Width="177" IsReadOnly="True">
					<TextBox.ContextMenu>
							<ContextMenu>
								<MenuItem x:Name='textBox_Compliance_Menu_OpenDeviceComplianceInBrowser' Header='Open Device Compliance in browser'/>
							</ContextMenu>
						</TextBox.ContextMenu>
					</TextBox>
                <Label x:Name="isEncrypted_label" Content="isEncrypted" HorizontalAlignment="Left" Margin="520,25,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
                <TextBox x:Name="isEncrypted_textBox" HorizontalAlignment="Left" Height="23" Margin="600,28,0,0" Text="isEncrypted" VerticalAlignment="Top" Width="177" IsReadOnly="True"/>
                <Label x:Name="lastSync_label" Content="Last Sync" HorizontalAlignment="Left" Margin="520,49,0,0" VerticalAlignment="Top" Height="24" FontWeight="Bold"/>
                <TextBox x:Name="lastSync_textBox" ToolTipService.ShowDuration="60000" HorizontalAlignment="Left" Height="23" Margin="600,52,0,0" Text="Last Sync" VerticalAlignment="Top" Width="177" IsReadOnly="True"/>
                <Label x:Name="primaryUser_label" Content="Primary User" HorizontalAlignment="Left" Margin="520,73,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
				<TextBox x:Name="primaryUser_textBox" ToolTipService.ShowDuration="500000" HorizontalAlignment="Left" Height="23" Margin="600,76,0,0" Text="Primary User" VerticalAlignment="Top" Width="177" IsReadOnly="True">
					<TextBox.ContextMenu>
						<ContextMenu>
							<MenuItem x:Name='primaryUser_textBox_Menu_Copy' Header='Copy'/>
							<MenuItem x:Name='primaryUser_textBox_Menu_OpenPrimaryUserInBrowser' Header='Open Primary User in browser'/>
						</ContextMenu>
					</TextBox.ContextMenu>
					<TextBox.ToolTip>
						<StackPanel>
							<TextBlock x:Name='textBlock_primaryUser_textBox_ToolTip_UPN' FontWeight="Bold" FontSize="14" Margin="0,0,0,5"></TextBlock>
							<TextBlock FontSize="14" FontWeight="Bold">Basic info</TextBlock>
							<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
							<TextBlock FontSize="14" x:Name='textBlock_primaryUser_textBox_ToolTip_BasicInfo' FontFamily="Consolas"></TextBlock>
							<TextBlock/>
							<TextBlock FontSize="14" FontWeight="Bold">proxyAddresses</TextBlock>
							<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
							<TextBlock FontSize="14" x:Name='textBlock_primaryUser_textBox_ToolTip_proxyAddresses' FontFamily="Consolas"></TextBlock>
							<TextBlock/>
							<TextBlock FontSize="14" FontWeight="Bold">otherMails</TextBlock>
							<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
							<TextBlock FontSize="14" x:Name='textBlock_primaryUser_textBox_ToolTip_otherMails' FontFamily="Consolas"></TextBlock>
							<TextBlock/>
							<TextBlock FontSize="14" FontWeight="Bold">onPremisesAttributes</TextBlock>
							<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
							<TextBlock FontSize="14" x:Name='textBlock_primaryUser_textBox_ToolTip_onPremisesAttributes' FontFamily="Consolas"></TextBlock>
							<TextBlock/>
							<TextBlock FontSize="14" FontWeight="Bold">onPremisesExtensionAttributes</TextBlock>
							<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
							<TextBlock FontSize="14" x:Name='textBlock_primaryUser_textBox_ToolTip_onPremisesExtensionAttributes' FontFamily="Consolas"></TextBlock>
							<TextBlock/>
						</StackPanel>
					</TextBox.ToolTip>
				</TextBox>
                <Label x:Name="AutopilotGroup_label" Content="Windows Autopilot" HorizontalAlignment="Left" Margin="790,1,0,0" VerticalAlignment="Top" Height="24" FontWeight="Bold"/>
                <Label x:Name="AutopilotEnrolled_label" Content="enrolled" HorizontalAlignment="Left" Margin="790,25,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
				<TextBox x:Name="AutopilotEnrolled_textBox" ToolTipService.ShowDuration="60000" HorizontalAlignment="Left" Height="23" Margin="858,28,0,0" Text="" VerticalAlignment="Top" Width="177" IsReadOnly="True">
					<TextBox.ContextMenu>
						<ContextMenu>
							<MenuItem x:Name='AutopilotEnrolled_textBox_Menu_OpenAutopilotDeviceInBrowser' Header='Open Autopilot devices in browser and paste device serial to search'/>
						</ContextMenu>
					</TextBox.ContextMenu>
				</TextBox>
                <Label x:Name="AutopilotGroupTag_label" Content="GroupTag" HorizontalAlignment="Left" Margin="790,49,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
                <TextBox x:Name="AutopilotGroupTag_textBox" ToolTipService.ShowDuration="60000" HorizontalAlignment="Left" Height="23" Margin="858,52,0,0" Text="" VerticalAlignment="Top" Width="177" IsReadOnly="True"/>
                <Label x:Name="Label_EnrollmentRestrictions" Content="Enrollment Restrictions" HorizontalAlignment="Left" Margin="1045,25,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
                <TextBox x:Name="textBox_EnrollmentRestrictions" ToolTipService.ShowDuration="60000" HorizontalAlignment="Left" Height="23" Margin="1190,28,0,0" Text="" VerticalAlignment="Top" Width="300" IsReadOnly="True">
					<TextBox.ContextMenu>
						<ContextMenu>
							<MenuItem x:Name='textBox_EnrollmentRestrictions_Menu_OpenEnrollmentRestrictionProfilesInBrowser' Header='Open Enrollment Restrictions Profiles Page in browser'/>
						</ContextMenu>
					</TextBox.ContextMenu>
				</TextBox>
				<Label x:Name="Label_EnrollmentStatusPage" Content="Enrollment Status Page" HorizontalAlignment="Left" Margin="1045,49,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>				
				<TextBox x:Name="textBox_EnrollmentStatusPage" ToolTipService.ShowDuration="60000" HorizontalAlignment="Left" Height="23" Margin="1190,52,0,0" Text="" VerticalAlignment="Top" Width="300" IsReadOnly="True">
					<TextBox.ContextMenu>
						<ContextMenu>
							<MenuItem x:Name='textBox_EnrollmentStatusPage_Menu_OpenESPProfileInBrowser' Header='Open Enrollment Status Page in browser'/>
						</ContextMenu>
					</TextBox.ContextMenu>
				</TextBox>
				<Label x:Name="AutopilotProfile_label" Content="Autopilot Depl. Profile" HorizontalAlignment="Left" Margin="1045,73,0,0" VerticalAlignment="Top" Height="26" FontWeight="Bold"/>
				<TextBox x:Name="AutopilotProfile_textBox" ToolTipService.ShowDuration="60000" HorizontalAlignment="Left" Height="23" Margin="1190,76,0,0" Text="" VerticalAlignment="Top" Width="300" IsReadOnly="True">
					<TextBox.ContextMenu>
						<ContextMenu>
							<MenuItem x:Name='AutopilotProfile_textBox_Menu_OpenAutopilotDeploymentProfileInBrowser' Header='Open Autopilot Deployment Profile in browser'/>
						</ContextMenu>
					</TextBox.ContextMenu>
				</TextBox>
            </Grid>
        </Border>
        <Border x:Name="IntuneDeviceDetailsBorderSignInUsers" Grid.Row="2" Grid.Column="0" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="25" MinHeight="25"/>
                    <RowDefinition Height="*" MinHeight="300"/>
                </Grid.RowDefinitions>
                <Label x:Name="IntuneDeviceDetails_RecentCheckins_label" Grid.Row="0" Content="Recent check-ins" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
				<TextBox x:Name="IntuneDeviceDetails_RecentCheckins_textBox" Grid.Row="1" Margin="5,5,5,5" TextWrapping="Wrap" Text="" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" ToolTipService.ShowDuration="500000">
					<TextBox.ContextMenu>
						<ContextMenu>
							<MenuItem x:Name='Latest_CheckIn_User_Menu_Copy' Header='Copy latest Checked-In user UPN'/>
							<MenuItem x:Name='Latest_CheckIn_User_Menu_Copy_Menu_OpenLatestCheckInUserInBrowser' Header='Open Latest Checked-In User in browser'/>
						</ContextMenu>
					</TextBox.ContextMenu>
					<TextBox.ToolTip>
						<StackPanel>
							<TextBlock FontWeight="Bold" FontSize="14" Margin="0,0,0,5">Latest or selected Checked-In user</TextBlock>
							<TextBlock x:Name='textBlock_LatestCheckInUser_textBox_ToolTip_UPN' FontWeight="Bold" FontSize="14" Margin="0,0,0,5"></TextBlock>
							<TextBlock FontSize="14" FontWeight="Bold">Basic info</TextBlock>
							<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
							<TextBlock FontSize="14" x:Name='textBlock_LatestCheckInUser_textBox_ToolTip_BasicInfo' FontFamily="Consolas"></TextBlock>
							<TextBlock/>
							<TextBlock FontSize="14" FontWeight="Bold">proxyAddresses</TextBlock>
							<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
							<TextBlock FontSize="14" x:Name='textBlock_LatestCheckInUser_textBox_ToolTip_proxyAddresses' FontFamily="Consolas"></TextBlock>
							<TextBlock/>
							<TextBlock FontSize="14" FontWeight="Bold">otherMails</TextBlock>
							<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
							<TextBlock FontSize="14" x:Name='textBlock_LatestCheckInUser_textBox_ToolTip_otherMails' FontFamily="Consolas"></TextBlock>
							<TextBlock/>
							<TextBlock FontSize="14" FontWeight="Bold">onPremisesAttributes</TextBlock>
							<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
							<TextBlock FontSize="14" x:Name='textBlock_LatestCheckInUser_textBox_ToolTip_onPremisesAttributes' FontFamily="Consolas"></TextBlock>
							<TextBlock/>
							<TextBlock FontSize="14" FontWeight="Bold">onPremisesExtensionAttributes</TextBlock>
							<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
							<TextBlock FontSize="14" x:Name='textBlock_LatestCheckInUser_textBox_ToolTip_onPremisesExtensionAttributes' FontFamily="Consolas"></TextBlock>
							<TextBlock/>
						</StackPanel>
					</TextBox.ToolTip>
				</TextBox>
            </Grid>
        </Border>
		<GridSplitter Grid.Row="2" Grid.Column="1" Width="5" HorizontalAlignment="Stretch" />
        <Border x:Name="IntuneDeviceDetailsBorderGroupMemberships" Grid.Row="2" Grid.Column="2" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid>
				<TabControl x:Name="tabControl_Device_and_User_GroupMemberships" Margin="5,5,0,0" Background="#FFE5EEFF" Foreground="White">
				   <TabControl.Resources>
					  <Style TargetType="TextBlock" x:Key="HeaderTextBlockStyle">
						 <Style.Triggers>
							<DataTrigger Binding="{Binding IsSelected, RelativeSource={RelativeSource AncestorType=TabItem}}" Value="True">
								<Setter Property="FontWeight" Value="Bold"/>
							 </DataTrigger>
						  </Style.Triggers>
					   </Style>
					</TabControl.Resources>
					<TabItem x:Name="TabItem_Device_GroupMemberships" Background="#FFE5EEFF">
						<TabItem.Header>
							<Grid ToolTipService.ShowDuration="500000">
								<Grid.ToolTip>
									<StackPanel>
										<TextBlock FontSize="14" FontWeight="Bold">Device Group Memberships</TextBlock>
										<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
										<TextBlock x:Name="TextBlock_TabItem_Device_GroupMemberships_Header_ToolTip" Text="" FontFamily="Consolas" FontSize="14"/>
									</StackPanel>
								</Grid.ToolTip>
								<TextBlock x:Name="TabItem_Device_GroupMemberships_Header" Text="Device Group Memberships" Style="{StaticResource HeaderTextBlockStyle}"/>
							</Grid>
						</TabItem.Header>
						<Grid x:Name="GridTabItem_Device_GroupMembershipsTAB">
							<ListView x:Name="listView_Device_GroupMemberships" Margin="5,5,5,5" IsManipulationEnabled="True">
								<!-- This makes our colored cells to fill whole cell background, not just text background -->
								<ListView.ItemContainerStyle>
									<Style TargetType="ListViewItem">
										<Setter Property="HorizontalContentAlignment" Value="Stretch"/>
									</Style>
								</ListView.ItemContainerStyle>
								<ListView.ContextMenu>
									<ContextMenu IsTextSearchEnabled="True">
										<MenuItem x:Name = 'ListView_GridTabItem_Device_GroupMembershipsTAB_Menu_Copy_DynamicRules' Header = 'Copy Dynamic Group rules to clipboard'/>
										<MenuItem x:Name = 'ListView_GridTabItem_Device_GroupMembershipsTAB_Menu_Copy_JSON' Header = 'Copy Azure AD Group JSON to clipboard'/>
										<Separator />
										<MenuItem x:Name = 'ListView_GridTabItem_Device_GroupMembershipsTAB_Menu_Open_Group_In_Browser' Header = 'Open Azure AD Group in browser'/>
									</ContextMenu>
								</ListView.ContextMenu>
								<ListView.View>
									<GridView>
										<GridViewColumn Width="350">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">DisplayName</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock FontWeight="Bold" Text="{Binding 'displayName'}" ToolTip="{Binding Path=description}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="50">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">Devices</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock HorizontalAlignment="Right" Text="{Binding 'YodamiittiCustomGroupMembersCountDevices'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="40">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">Users</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock HorizontalAlignment="Right" Text="{Binding 'YodamiittiCustomGroupMembersCountUsers'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="75">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">GroupType</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock Text="{Binding 'YodamiittiCustomGroupType'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="50">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">SecurityEnabled</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock Text="{Binding 'securityEnabled'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="98">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">MembeshipType</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock Text="{Binding 'YodamiittiCustomMembershipType'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="Auto">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">MembeshipRule</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock Text="{Binding 'membershipRule'}" ToolTip="{Binding Path=membershipRule}" ToolTipService.ShowDuration="500000">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
									</GridView>
								</ListView.View>
							</ListView>
						</Grid>
					</TabItem>
					<TabItem x:Name="TabItem_User_GroupMembershipsTAB" Background="#FFE5EEFF">
						<TabItem.Header>
							<Grid ToolTipService.ShowDuration="500000">
								<Grid.ToolTip>
									<StackPanel>
										<TextBlock FontSize="14" FontWeight="Bold">Primary User Group Memberships</TextBlock>
										<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
										<TextBlock x:Name="TextBlock_TabItem_User_GroupMembershipsTAB_Header_ToolTip" Text="" FontFamily="Consolas" FontSize="14"/>
									</StackPanel>
								</Grid.ToolTip>
								<TextBlock x:Name="TabItem_User_GroupMembershipsTAB_Header" Text="Primary User Group Memberships" Style="{StaticResource HeaderTextBlockStyle}"/>
							</Grid>
						</TabItem.Header>
						<Grid>
							<ListView x:Name="listView_PrimaryUser_GroupMemberships" Margin="5,5,5,5" IsManipulationEnabled="True">
								<!-- This makes our colored cells to fill whole cell background, not just text background -->
								<ListView.ItemContainerStyle>
									<Style TargetType="ListViewItem">
										<Setter Property="HorizontalContentAlignment" Value="Stretch"/>
									</Style>
								</ListView.ItemContainerStyle>
								<ListView.ContextMenu>
									<ContextMenu IsTextSearchEnabled="True">
										<MenuItem x:Name = 'ListView_GridTabItem_PrimaryUser_GroupMembershipsTAB_Menu_Copy_DynamicRules' Header = 'Copy Dynamic Group rules to clipboard'/>
										<MenuItem x:Name = 'ListView_GridTabItem_PrimaryUser_GroupMembershipsTAB_Menu_Copy_JSON' Header = 'Copy Azure AD Group JSON to clipboard'/>
										<Separator />
										<MenuItem x:Name = 'ListView_GridTabItem_PrimaryUser_GroupMembershipsTAB_Menu_Open_Group_In_Browser' Header = 'Open Azure AD Group in browser'/>
									</ContextMenu>
								</ListView.ContextMenu>
								<ListView.View>
									<GridView>
										<GridViewColumn Width="350">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">DisplayName</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="0,0,0,0">
														<Grid.Style>
															<Style TargetType="{x:Type Grid}">
																<Style.Triggers>
																	<DataTrigger Binding="{Binding YodamiittiCustomGroupType}" Value="DirectoryRole">
																		<Setter Property="Background" Value="yellow"/>
																	</DataTrigger>
																	<DataTrigger Binding="{Binding displayName}" Value="Global Administrator">
																		<Setter Property="Background" Value="red"/>
																	</DataTrigger>
																</Style.Triggers>
															</Style>
														</Grid.Style>
														<TextBlock FontWeight="Bold" HorizontalAlignment="Left" Text="{Binding displayName}" ToolTip="{Binding Path=description}" Padding="0" Margin="0" />
													</Grid>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="50">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">Devices</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock HorizontalAlignment="Right" Text="{Binding 'YodamiittiCustomGroupMembersCountDevices'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="40">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">Users</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock HorizontalAlignment="Right" Text="{Binding 'YodamiittiCustomGroupMembersCountUsers'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="75">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">GroupType</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="0,0,0,0">
														<Grid.Style>
															<Style TargetType="{x:Type Grid}">
																<Style.Triggers>
																	<DataTrigger Binding="{Binding YodamiittiCustomGroupType}" Value="DirectoryRole">
																		<Setter Property="Background" Value="yellow"/>
																	</DataTrigger>
																	<DataTrigger Binding="{Binding displayName}" Value="Global Administrator">
																		<Setter Property="Background" Value="red"/>
																	</DataTrigger>
																</Style.Triggers>
															</Style>
														</Grid.Style>
														<TextBlock HorizontalAlignment="Left" Text="{Binding YodamiittiCustomGroupType}" Padding="0" Margin="0" />
													</Grid>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="50">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">SecurityEnabled</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock Text="{Binding 'securityEnabled'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="98">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">MembeshipType</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock Text="{Binding 'YodamiittiCustomMembershipType'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="Auto">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">MembeshipRule</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock Text="{Binding 'membershipRule'}" ToolTip="{Binding Path=membershipRule}" ToolTipService.ShowDuration="500000">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
									</GridView>
								</ListView.View>
							</ListView>
						</Grid>
					</TabItem>
					<TabItem x:Name="TabItem_LatestCheckedInUser_GroupMembershipsTAB" Background="#FFE5EEFF">
						<TabItem.Header>
							<Grid ToolTipService.ShowDuration="500000">
								<Grid.ToolTip>
									<StackPanel>
										<TextBlock FontSize="14" FontWeight="Bold">Checked-in User's Group Memberships</TextBlock>
										<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="0,0" />
										<TextBlock x:Name="TextBlock_TabItem_LatestCheckedInUser_GroupMembershipsTAB_Header_ToolTip" Text="" FontFamily="Consolas" FontSize="14"/>
									</StackPanel>
								</Grid.ToolTip>
								<TextBlock x:Name="TabItem_LatestCheckedInUser_GroupMembershipsTAB_Header" Text="Checked-In User Group Memberships" Style="{StaticResource HeaderTextBlockStyle}"/>
							</Grid>
						</TabItem.Header>
						<Grid>
							<ListView x:Name="listView_LatestCheckedInUser_GroupMemberships" Margin="5,5,5,5" IsManipulationEnabled="True">
								<!-- This makes our colored cells to fill whole cell background, not just text background -->
								<ListView.ItemContainerStyle>
									<Style TargetType="ListViewItem">
										<Setter Property="HorizontalContentAlignment" Value="Stretch"/>
									</Style>
								</ListView.ItemContainerStyle>
								<ListView.ContextMenu>
									<ContextMenu IsTextSearchEnabled="True">
										<MenuItem x:Name = 'ListView_GridTabItem_LatestCheckedInUser_GroupMembershipsTAB_Menu_Copy_DynamicRules' Header = 'Copy Dynamic Group rules to clipboard'/>
										<MenuItem x:Name = 'ListView_GridTabItem_LatestCheckedInUser_GroupMembershipsTAB_Menu_Copy_JSON' Header = 'Copy Azure AD Group JSON to clipboard'/>
										<Separator />
										<MenuItem x:Name = 'ListView_GridTabItem_LatestCheckedInUser_GroupMembershipsTAB_Menu_Open_Group_In_Browser' Header = 'Open Azure AD Group in browser'/>
									</ContextMenu>
								</ListView.ContextMenu>
								<ListView.View>
									<GridView>
										<GridViewColumn Width="350">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">DisplayName</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="0,0,0,0">
														<Grid.Style>
															<Style TargetType="{x:Type Grid}">
																<Style.Triggers>
																	<DataTrigger Binding="{Binding YodamiittiCustomGroupType}" Value="DirectoryRole">
																		<Setter Property="Background" Value="yellow"/>
																	</DataTrigger>
																	<DataTrigger Binding="{Binding displayName}" Value="Global Administrator">
																		<Setter Property="Background" Value="red"/>
																	</DataTrigger>
																</Style.Triggers>
															</Style>
														</Grid.Style>
														<TextBlock FontWeight="Bold" HorizontalAlignment="Left" Text="{Binding displayName}" ToolTip="{Binding Path=description}" Padding="0" Margin="0" />
													</Grid>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="50">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">Devices</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock HorizontalAlignment="Right" Text="{Binding 'YodamiittiCustomGroupMembersCountDevices'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="40">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">Users</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock HorizontalAlignment="Right" Text="{Binding 'YodamiittiCustomGroupMembersCountUsers'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="75">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">GroupType</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="0,0,0,0">
														<Grid.Style>
															<Style TargetType="{x:Type Grid}">
																<Style.Triggers>
																	<DataTrigger Binding="{Binding YodamiittiCustomGroupType}" Value="DirectoryRole">
																		<Setter Property="Background" Value="yellow"/>
																	</DataTrigger>
																	<DataTrigger Binding="{Binding displayName}" Value="Global Administrator">
																		<Setter Property="Background" Value="red"/>
																	</DataTrigger>
																</Style.Triggers>
															</Style>
														</Grid.Style>
														<TextBlock HorizontalAlignment="Left" Text="{Binding YodamiittiCustomGroupType}" Padding="0" Margin="0" />
													</Grid>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="50">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">SecurityEnabled</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock Text="{Binding 'securityEnabled'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="98">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">MembeshipType</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock Text="{Binding 'YodamiittiCustomMembershipType'}">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
										<GridViewColumn Width="Auto">
											<GridViewColumn.Header>
												<GridViewColumnHeader FontWeight="Bold">MembeshipRule</GridViewColumnHeader>
											</GridViewColumn.Header>
											<GridViewColumn.CellTemplate>
												<DataTemplate>
													<TextBlock Text="{Binding 'membershipRule'}" ToolTip="{Binding Path=membershipRule}" ToolTipService.ShowDuration="500000">
													</TextBlock>
												</DataTemplate>
											</GridViewColumn.CellTemplate>
										</GridViewColumn>
									</GridView>
								</ListView.View>
							</ListView>
						</Grid>
					</TabItem>
				</TabControl>
            </Grid>
        </Border>
        <GridSplitter Grid.Row="2" Grid.Column="3" Grid.RowSpan="5" Width="5" HorizontalAlignment="Stretch" />
		<Border x:Name="IntuneDeviceDetailsBorderXAML" Grid.Row="2" Grid.Column="4" Grid.RowSpan="5" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid>
				<TabControl x:Name="tabControlDetailsXAML" Margin="5,5,0,0" Background="#FFE5EEFF" Foreground="White">
					<TabControl.Resources>
					  <Style TargetType="TextBlock" x:Key="HeaderTextBlockStyle">
						 <Style.Triggers>
							<DataTrigger Binding="{Binding IsSelected, RelativeSource={RelativeSource AncestorType=TabItem}}" Value="True">
								<Setter Property="FontWeight" Value="Bold"/>
							 </DataTrigger>
						  </Style.Triggers>
					   </Style>
					</TabControl.Resources>
					<TabItem x:Name="TabItem_Overview" Background="#FFE5EEFF">
						<TabItem.Header>
							<TextBlock Text="Overview" Style="{StaticResource HeaderTextBlockStyle}" FontSize="16" FontWeight="Bold"/>
						</TabItem.Header>
						<Grid x:Name="GridTabItem_Overview_TAB">
							<Border x:Name="BorderGridTabItem_Overview_TAB" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" Background="#FFF7F7F7" CornerRadius="8">
								<Grid x:Name="GridTabItem_Overview_TABWindow">
									<ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" >
										<StackPanel>
											<TextBlock x:Name="TextBlock_Overview_DeviceName" FontSize="16" FontWeight="Bold" Margin="6,0">DeviceName</TextBlock>
											<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="6,0" />
											<TextBox x:Name="TextBox_Overview_Device" TextWrapping="NoWrap" Text="" Padding="5" VerticalScrollBarVisibility="Hidden" HorizontalScrollBarVisibility="Hidden" Margin="5" IsReadOnly="True" FontFamily="Consolas" FontSize="14"/>
											<TextBlock/>
											<TextBlock x:Name="TextBlock_Overview_PrimaryUserName" FontSize="16" FontWeight="Bold" Margin="6,0">PrimaryUser</TextBlock>
											<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="6,0" />
											<TextBox x:Name="TextBox_Overview_PrimaryUser" TextWrapping="NoWrap" Text="" Padding="5" VerticalScrollBarVisibility="Hidden" HorizontalScrollBarVisibility="Hidden" Margin="5" IsReadOnly="True" FontFamily="Consolas" FontSize="14"/>
											<TextBlock/>
											<TextBlock x:Name="TextBlock_Overview_LatestCheckedInUserName" FontSize="16" FontWeight="Bold" Margin="6,0">Latest or Selected Checked-In User</TextBlock>
											<Border BorderBrush="Silver" BorderThickness="0,1,0,0" Margin="6,0" />
											<TextBox x:Name="TextBox_Overview_LatestCheckedInUser" TextWrapping="NoWrap" Text="" Padding="5" VerticalScrollBarVisibility="Hidden" HorizontalScrollBarVisibility="Hidden" Margin="5" IsReadOnly="True" FontFamily="Consolas" FontSize="14"/>
										</StackPanel>
									</ScrollViewer>
								</Grid>
							</Border>
						</Grid>
					</TabItem>
					<TabItem x:Name="IntuneDeviceJSON" Background="#FFE5EEFF">
						<TabItem.Header>
							<TextBlock Text="Intune Device JSON" Style="{StaticResource HeaderTextBlockStyle}"/>
						</TabItem.Header>
						<Grid x:Name="GridIntuneDeviceJSONTAB">
							<Border x:Name="BorderGridIntuneDeviceJSONTAB" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" Background="#FFF7F7F7" CornerRadius="8">
								<Grid x:Name="GridIntuneDeviceJSONTABWindow">
									<TextBox x:Name="IntuneDeviceDetails_json_textBox" TextWrapping="Wrap" Text="" Padding="5" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Margin="5" IsReadOnly="True" />
								</Grid>
							</Border>
						</Grid>
					</TabItem>
					<TabItem x:Name="AzureADDeviceJSON" Background="#FFE5EEFF">
						<TabItem.Header>
							<TextBlock Text="Azure AD Device JSON" Style="{StaticResource HeaderTextBlockStyle}"/>
						</TabItem.Header>
						<Grid x:Name="GridAzureADDeviceJSONTAB">
							<Border x:Name="BorderGridAzureADDeviceJSONTAB" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" Background="#FFF7F7F7" CornerRadius="8">
								<Grid x:Name="GridAzureADDeviceJSONTABWindow">
									<TextBox x:Name="AzureADDeviceDetails_json_textBox" TextWrapping="Wrap" Text="" Padding="5" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Margin="5" IsReadOnly="True" />
								</Grid>
							</Border>
						</Grid>
					</TabItem>
					<TabItem x:Name="PrimaryUserJSON" Background="#FFE5EEFF">
						<TabItem.Header>
							<TextBlock Text="Primary User JSON" Style="{StaticResource HeaderTextBlockStyle}"/>
						</TabItem.Header>
						<Grid x:Name="GridPrimaryUserJSONTAB">
							<Border x:Name="BorderGridPrimaryUserJSONTAB" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" Background="#FFF7F7F7" CornerRadius="8">
								<Grid x:Name="GridPrimaryUserJSONTABWindow">
									<TextBox x:Name="PrimaryUserDetails_json_textBox" TextWrapping="Wrap" Text="" Padding="5" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Margin="5" IsReadOnly="True" />
								</Grid>
							</Border>
						</Grid>
					</TabItem>
					<TabItem x:Name="LatestCheckedInUserJSON" Background="#FFE5EEFF">
						<TabItem.Header>
							<TextBlock Text="Latest Checked-In User JSON" Style="{StaticResource HeaderTextBlockStyle}"/>
						</TabItem.Header>
						<Grid x:Name="GridLatestCheckedInUserJSONTAB">
							<Border x:Name="BorderGridLatestCheckedInUserJSONTAB" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" Background="#FFF7F7F7" CornerRadius="8">
								<Grid x:Name="GridLatestCheckedInUserJSONTABWindow">
									<TextBox x:Name="LatestCheckedInUser_json_textBox" TextWrapping="Wrap" Text="" Padding="5" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Margin="5" IsReadOnly="True" />
								</Grid>
							</Border>
						</Grid>
					</TabItem>
					<TabItem x:Name="AutopilotDeviceJSON" Background="#FFE5EEFF">
						<TabItem.Header>
							<TextBlock Text="Autopilot Device JSON" Style="{StaticResource HeaderTextBlockStyle}"/>
						</TabItem.Header>
						<Grid x:Name="GridAutopilotDeviceJSONTAB">
							<Border x:Name="BorderGridAutopilotDeviceJSONTAB" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" Background="#FFF7F7F7" CornerRadius="8">
								<Grid x:Name="GridAutopilotDeviceJSONTABWindow">
									<TextBox x:Name="AutopilotDevice_json_textBox" TextWrapping="Wrap" Text="" Padding="5" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Margin="5" IsReadOnly="True" />
								</Grid>
							</Border>
						</Grid>
					</TabItem>
					<TabItem x:Name="AutopilotEnrollmentProfileJSON" Background="#FFE5EEFF">
						<TabItem.Header>
							<TextBlock Text="Autopilot Enrollment Profile JSON" Style="{StaticResource HeaderTextBlockStyle}"/>
						</TabItem.Header>
						<Grid x:Name="GridAutopilotEnrollmentProfileJSONTAB">
							<Border x:Name="BorderGridAutopilotEnrollmentProfileJSONTAB" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" Background="#FFF7F7F7" CornerRadius="8">
								<Grid x:Name="GridAutopilotEnrollmentProfileJSONTABWindow">
									<TextBox x:Name="AutopilotEnrollmentProfile_json_textBox" TextWrapping="Wrap" Text="" Padding="5" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Margin="5" IsReadOnly="True" />
								</Grid>
							</Border>
						</Grid>
					</TabItem>
					<TabItem x:Name="ApplicationJSON" Background="#FFE5EEFF">
						<TabItem.Header>
							<TextBlock Text="Selected Application JSON" Style="{StaticResource HeaderTextBlockStyle}"/>
						</TabItem.Header>
						<Grid x:Name="GridApplicationJSONTAB">
							<Border x:Name="BorderGridApplicationJSONTAB" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" Background="#FFF7F7F7" CornerRadius="8">
								<Grid x:Name="GridApplicationJSONTABWindow">
									<TextBox x:Name="Application_json_textBox" TextWrapping="Wrap" Text="" Padding="5" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Margin="5" IsReadOnly="True" />
								</Grid>
							</Border>
						</Grid>
					</TabItem>
					<TabItem x:Name="AzureADGroupJSON" Background="#FFE5EEFF">
						<TabItem.Header>
							<TextBlock Text="Selected Azure AD Group JSON" Style="{StaticResource HeaderTextBlockStyle}"/>
						</TabItem.Header>
						<Grid x:Name="AzureADGroupJSONTAB">
							<Border x:Name="BorderGridAzureADGroupJSONTAB" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" Background="#FFF7F7F7" CornerRadius="8">
								<Grid x:Name="GridAzureADGroupJSONTABWindow">
									<TextBox x:Name="AzureAD_Group_json_textBox" TextWrapping="Wrap" Text="" Padding="5" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Margin="5" IsReadOnly="True" />
								</Grid>
							</Border>
						</Grid>
					</TabItem>
					<TabItem x:Name="ConfigurationJSON" Background="#FFE5EEFF">
						<TabItem.Header>
							<TextBlock Text="Selected Configuration JSON" Style="{StaticResource HeaderTextBlockStyle}"/>
						</TabItem.Header>
						<Grid x:Name="GridConfigurationJSONTAB">
							<Border x:Name="BorderGridConfigurationJSONTAB" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" Background="#FFF7F7F7" CornerRadius="8">
								<Grid x:Name="GridConfigurationJSONTABWindow">
									<TextBox x:Name="Configuration_json_textBox" TextWrapping="Wrap" Text="" Padding="5" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Margin="5" IsReadOnly="True" />
								</Grid>
							</Border>
						</Grid>
					</TabItem>
				</TabControl>
            </Grid>
        </Border>
		<GridSplitter Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="3" Height="5" HorizontalAlignment="Stretch" />
        <Border x:Name="IntuneDeviceDetailsBorderApplications" Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="3" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="25" MinHeight="25"/>
                    <RowDefinition Height="*" MinHeight="150"/>
                </Grid.RowDefinitions>
                <Label x:Name="IntuneDeviceDetails_ApplicationAssignments_label" Grid.Row="0" Content="Application Assignments" Height="27" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
				<Label x:Name="IntuneDeviceDetails_ApplicationAssignments_SelectUser_label" Grid.Row="0" Content="Select user" Height="27" HorizontalAlignment="Right" Margin="0,0,310,0" VerticalAlignment="Top" FontWeight="Bold"/>
				<ComboBox x:Name="IntuneDeviceDetails_ApplicationAssignments_SelectUser_ComboBox" Grid.Row="0" HorizontalAlignment="Right" Margin="0,5,6,0" VerticalAlignment="Top" Width="300" IsEditable="False" FontFamily="Consolas" FontSize="14">
				</ComboBox>
				<ListView x:Name="listView_ApplicationAssignments" Grid.Row="1" Margin="5,5,5,5" IsManipulationEnabled="True">
                    <!-- This makes our colored cells to fill whole cell background, not just text background -->
                    <ListView.ItemContainerStyle>
                        <Style TargetType="ListViewItem">
                            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                        </Style>
                    </ListView.ItemContainerStyle>
					<ListView.ContextMenu>
						<ContextMenu IsTextSearchEnabled="True">
							<!-- <MenuItem x:Name = 'listView_ApplicationAssignments_Menu_Copy_ApplicationBasicInfo' Header = 'Copy Application basic info to clipboard'/> -->
							<MenuItem x:Name = 'listView_ApplicationAssignments_Menu_Copy_JSON' Header = 'Copy Application JSON to clipboard'/>
							<MenuItem x:Name = 'listView_ApplicationAssignments_Menu_Copy_DetectionRules_Powershell_to_Clipboard' Header = 'Copy Win32 Application Detection Rules Powershell script(s) to clipboard'/>
							<MenuItem x:Name = 'listView_ApplicationAssignments_Menu_Copy_requirementRules_Powershell_to_Clipboard' Header = 'Copy Win32 Application Requirement Rules Powershell script(s) to clipboard'/>
							<Separator />
							<MenuItem x:Name = 'listView_ApplicationAssignments_Menu_Open_Application_In_Browser' Header = 'Open Application in browser'/>
							<MenuItem x:Name = 'listView_ApplicationAssignments_Menu_Open_ApplicationAssignmentBroup_In_Browser' Header = 'Open Assignment Group in browser'/>
							<MenuItem x:Name = 'listView_ApplicationAssignments_Menu_Open_ApplicationAssignmentFilter_In_Browser' Header = 'Open Filter in browser'/>
						</ContextMenu>
					</ListView.ContextMenu>
                    <ListView.View>
                        <GridView>
                            <GridViewColumn Width="70">
                                <GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">context</GridViewColumnHeader>
								</GridViewColumn.Header>
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
                                            <TextBlock HorizontalAlignment="Center" Text="{Binding context}" ToolTip="{Binding Path=contextToolTip}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
							<GridViewColumn Width="175">
                                <GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">Application type</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock Text="{Binding 'odatatype'}">
                                        </TextBlock>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
							<GridViewColumn Width="300">
                                <GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">displayName</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock FontWeight="Bold" Text="{Binding displayName}" ToolTip="{Binding Path=displayNameToolTip}">
                                        </TextBlock>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
							<GridViewColumn Width="70">
                                <GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">Version</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock Text="{Binding version}">
                                        </TextBlock>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
                            <GridViewColumn Width="80">
                                <GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">Intent</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock HorizontalAlignment="Center" Text="{Binding assignmentIntent}" />
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
							<GridViewColumn Width="90">
								<GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">IncludeExclude</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
                                            <Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
                                                        <DataTrigger Binding="{Binding IncludeExclude}" Value="Excluded">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </Grid.Style>
                                            <TextBlock HorizontalAlignment="Center" Text="{Binding IncludeExclude}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
                            <GridViewColumn Width="130">
								<GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">installState</GridViewColumnHeader>
								</GridViewColumn.Header>
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
														<DataTrigger Binding="{Binding installState}" Value="notApplicable">
                                                            <Setter Property="Background" Value="yellow"/>
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
							<GridViewColumn Width="70">
                                <GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">GroupType</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock Text="{Binding YodamiittiCustomMembershipType}">
                                        </TextBlock>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>							
                            <GridViewColumn Width="250">
								<GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">assignmentGroup</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
                                            <Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
														<DataTrigger Binding="{Binding assignmentGroup}" Value="Application does not have any assignments!">
															<Setter Property="Background" Value="#FF6347"/>
														</DataTrigger>
														<DataTrigger Binding="{Binding installState}" Value="notApplicable">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
														<DataTrigger Binding="{Binding IncludeExclude}" Value="Excluded">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
														<DataTrigger Binding="{Binding assignmentGroup}" Value="unknown (possible nested group or removed assignment)">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </Grid.Style>
                                            <TextBlock HorizontalAlignment="Left" Text="{Binding assignmentGroup}" ToolTip="{Binding Path=AssignmentGroupToolTip}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
							</GridViewColumn>
							<GridViewColumn Width="105">
                                <GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">Group Members</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock HorizontalAlignment="Right" Text="{Binding YodamiittiCustomGroupMembers}">
                                        </TextBlock>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>							
							<GridViewColumn Width="250">
								<GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">Filter</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
											<Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
														<DataTrigger Binding="{Binding installState}" Value="notApplicable">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </Grid.Style>
                                            <TextBlock HorizontalAlignment="Left" Text="{Binding filter}" ToolTip="{Binding Path=filterToolTip}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
							</GridViewColumn>
							<GridViewColumn Width="80">
                                <GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">FilterMode</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
									<Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
											<Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
														<DataTrigger Binding="{Binding installState}" Value="notApplicable">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </Grid.Style>
                                            <TextBlock HorizontalAlignment="Left" Text="{Binding filterMode}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>							
                        </GridView>
                    </ListView.View>
                </ListView>
            </Grid>
        </Border>
		<GridSplitter Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="3" Height="5" HorizontalAlignment="Stretch" />
        <Border x:Name="IntuneDeviceDetailsBorderConfigurations" Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="3" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="25" MinHeight="25"/>
                    <RowDefinition Height="*" MinHeight="150"/>
                </Grid.RowDefinitions>
                <Label x:Name="IntuneDeviceDetails_ConfigurationsAssignments_label" Grid.Row="0" Content="Configurations Assignments" Height="27" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
                <ListView x:Name="listView_ConfigurationsAssignments" Grid.Row="1" Margin="5,5,5,5" IsManipulationEnabled="True">
                    <!-- This makes our colored cells to fill whole cell background, not just text background -->
                    <ListView.ItemContainerStyle>
                        <Style TargetType="ListViewItem">
                            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                        </Style>
                    </ListView.ItemContainerStyle>
					<ListView.ContextMenu>
						<ContextMenu IsTextSearchEnabled="True">
							<!-- <MenuItem x:Name = 'listView_ConfigurationsAssignments_Menu_Copy_ConfigurationBasicInfo' Header = 'Copy Configuration basic info to clipboard'/> -->
							<MenuItem x:Name = 'listView_ConfigurationsAssignments_Menu_Copy_JSON' Header = 'Copy Configuration Profile JSON to clipboard'/>
							<Separator />
							<!-- <MenuItem x:Name = 'listView_ConfigurationsAssignments_Menu_Open_Configuration_In_Browser' Header = 'Open Configuration in browser'/> -->
							<MenuItem x:Name = 'listView_ConfigurationsAssignments_Menu_Open_ConfigurationAssignmentBroup_In_Browser' Header = 'Open Assignment Group in browser'/>
							<MenuItem x:Name = 'listView_ConfigurationsAssignments_Menu_Open_ConfigurationAssignmentFilter_In_Browser' Header = 'Open Filter in browser'/>
						</ContextMenu>
					</ListView.ContextMenu>
                    <ListView.View>
                        <GridView>
                            <GridViewColumn Width="70">
								<GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">context</GridViewColumnHeader>
								</GridViewColumn.Header>
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
                                            <TextBlock HorizontalAlignment="Center" Text="{Binding context}" ToolTip="{Binding Path=contextToolTip}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
							<GridViewColumn Width="175">
                                <GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">Configuration type</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock Text="{Binding 'odatatype'}" />
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
                            <GridViewColumn Width="300">
								<GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">displayName</GridViewColumnHeader>
								</GridViewColumn.Header>
                                <GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock FontWeight="Bold" Text="{Binding displayName}" ToolTip="{Binding Path=displayNameToolTip}">
                                        </TextBlock>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
							<GridViewColumn Width="190">
								<GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">userPrincipalName</GridViewColumnHeader>
								</GridViewColumn.Header>
                                <GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock FontWeight="Bold" Text="{Binding userPrincipalName}"/>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
							<GridViewColumn Width="90">
								<GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">IncludeExclude</GridViewColumnHeader>
								</GridViewColumn.Header>
                                <GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
                                            <Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
                                                        <DataTrigger Binding="{Binding IncludeExclude}" Value="Excluded">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </Grid.Style>
                                            <TextBlock HorizontalAlignment="Center" Text="{Binding IncludeExclude}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
							<GridViewColumn Width="85">
								<GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">state</GridViewColumnHeader>
								</GridViewColumn.Header>
                                <GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
                                            <Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
                                                        <DataTrigger Binding="{Binding state}" Value="Succeeded">
                                                            <Setter Property="Background" Value="#7FFF00"/>
                                                        </DataTrigger>
														<DataTrigger Binding="{Binding state}" Value="Conflict">
                                                            <Setter Property="Background" Value="#FF6347"/>
                                                        </DataTrigger>
														<DataTrigger Binding="{Binding state}" Value="Error">
                                                            <Setter Property="Background" Value="#FF6347"/>
                                                        </DataTrigger>
														<DataTrigger Binding="{Binding state}" Value="Not applicable">
                                                            <Setter Property="Background" Value="yellow"/>
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
							<GridViewColumn Width="70">
                                <GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">GroupType</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock Text="{Binding YodamiittiCustomMembershipType}">
                                        </TextBlock>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>							
							<GridViewColumn Width="250">
								<GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">assignmentGroup</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
                                            <Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
														<DataTrigger Binding="{Binding assignmentGroup}" Value="Policy does not have any assignments!">
															<Setter Property="Background" Value="#FF6347"/>
														</DataTrigger>
														<DataTrigger Binding="{Binding state}" Value="Not applicable">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
														<DataTrigger Binding="{Binding IncludeExclude}" Value="Excluded">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
														<DataTrigger Binding="{Binding assignmentGroup}" Value="unknown (possible nested group or removed assignment)">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </Grid.Style>
                                            <TextBlock HorizontalAlignment="Left" Text="{Binding assignmentGroup}" ToolTip="{Binding Path=AssignmentGroupToolTip}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
							</GridViewColumn>
							<GridViewColumn Width="105">
                                <GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">Group Members</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBlock HorizontalAlignment="Right" Text="{Binding YodamiittiCustomGroupMembers}">
                                        </TextBlock>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>							
							<GridViewColumn Width="250">
								<GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">Filter</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
											<Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
														<DataTrigger Binding="{Binding state}" Value="Not applicable">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </Grid.Style>
											<TextBlock HorizontalAlignment="Left" Text="{Binding filter}" ToolTip="{Binding Path=filterToolTip}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
							</GridViewColumn>
							<GridViewColumn Width="80">
                                <GridViewColumn.Header>
									<GridViewColumnHeader FontWeight="Bold">FilterMode</GridViewColumnHeader>
								</GridViewColumn.Header>
								<GridViewColumn.CellTemplate>
                                    <DataTemplate>
										<Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="-6,0,-6,0">
											<Grid.Style>
                                                <Style TargetType="{x:Type Grid}">
                                                    <Style.Triggers>
														<DataTrigger Binding="{Binding state}" Value="Not applicable">
                                                            <Setter Property="Background" Value="yellow"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </Grid.Style>
											<TextBlock HorizontalAlignment="Left" Text="{Binding filterMode}" Padding="0" Margin="0" />
                                        </Grid>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>							
                        </GridView>
                    </ListView.View>
                </ListView>
            </Grid>
        </Border>
        <Border x:Name="IntuneDeviceDetailsBorderBottom" Grid.Row="7" Grid.Column="0" Grid.ColumnSpan="5" BorderBrush="Black" BorderThickness="1" Margin="0,0,2,2" CornerRadius="8" Background="#FFF7F7F7">
            <Grid x:Name="IntuneDeviceDetailsGridBottom">
				<TextBox x:Name="AboutTAB_textBox_author" HorizontalAlignment="Right" Height="26" Margin="0,2,105,0" TextWrapping="Wrap" Text="Author: Petri.Paavola@yodamiitti.fi - Microsoft MVP" VerticalAlignment="Top" Width="360" FontSize="14" IsReadOnly="True" BorderThickness="0,0,0,0" IsReadOnlyCaretVisible="True" Background="Transparent" Focusable="False">
					<MenuItem.ToolTip>
						<StackPanel Orientation="Horizontal">
							<Label>Right click me</Label>
						</StackPanel>
					</MenuItem.ToolTip>
					<TextBox.ContextMenu>
						<ContextMenu>
							<MenuItem x:Name='AboutTAB_textBox_author_Menu_Copy' Header='Copy email address'/>
							<MenuItem x:Name='AboutTAB_textBox_author_Menu_OpenEmailAddress' Header='Open email address in default mail app'/>
						</ContextMenu>
					</TextBox.ContextMenu>
				</TextBox>				
				<TextBox x:Name="AboutTAB_textBox_github_link" HorizontalAlignment="Right" Height="26" Margin="0,20,105,0" TextWrapping="Wrap" Text="https://github.com/petripaavola/IntuneDeviceDetailsGUI" VerticalAlignment="Top" Width="360" FontSize="14" IsReadOnly="True" BorderThickness="0,0,0,0" IsReadOnlyCaretVisible="True" Background="Transparent" Focusable="False">
					<MenuItem.ToolTip>
						<StackPanel Orientation="Horizontal">
							<Label>Right click me</Label>
						</StackPanel>
					</MenuItem.ToolTip>
					<TextBox.ContextMenu>
						<ContextMenu>
							<MenuItem x:Name='AboutTAB_textBox_github_link_Menu_Copy' Header='Copy url'/>
							<MenuItem x:Name='AboutTAB_textBox_github_link_Menu_OpenDeviceInBrowser' Header='Open GitHub link in browser to check updates'/>
						</ContextMenu>
					</TextBox.ContextMenu>
				</TextBox>
                <Image x:Name="bottomRightLogoimage" HorizontalAlignment="Right" Height="46" Margin="0,0,9.8,0" VerticalAlignment="Top" Width="133">
					<MenuItem.ToolTip>
						<StackPanel Orientation="Horizontal">
							<Label>Right Click me to open yodamiitti.com website</Label>
						</StackPanel>
					</MenuItem.ToolTip>
					<Image.ContextMenu>
						<ContextMenu>
							<MenuItem x:Name='AboutTAB_image_Yodamiitti_Menu_VisitYodamiittiWebsite' Header='Open Yodamiitti.com website in browser'/>
						</ContextMenu>
					</Image.ContextMenu>
				</Image>
                <TextBox x:Name="bottom_textBox" HorizontalAlignment="Left" Height="26" Margin="10,10,0,0" Text="" VerticalAlignment="Top" Width="510" FontSize="16" IsReadOnly="True"/>
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
try {
	$Form = [Windows.Markup.XamlReader]::Load( $reader )
} catch {
	Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
	Exit 1
}

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

	Start-Sleep -Milliseconds $TimeOutBetweenGraphAPIRequests

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
		# and check that $Top= parameter was NOT used. With $Top= parameter we can limit search results
		# but that almost always results .nextLink being present if there is more data than top specified
		# If we specified $Top= ourselfes then we don't want to get nextLink values
		#
		# So get GraphAllPages if there is valid nextlink and not $Top= used in url originally
		if (($MSGraphRequest.'@odata.nextLink' -like 'https://*') -and (-not ($url.Contains('$top=')))) {

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


function Get-CheckedInUsersGroupMemberships {
    param (
        [Parameter(Mandatory = $false)]
        $SelectedUser=$false
    )

	if($SelectedUser) {
		Write-Verbose "Get group memberships for user $($SelectedUser.UserPrincipalName)"
	} else {
		Write-Verbose "Get group memberships for latest signed-in user"
	}


	# Get Logged on users information
    [String]$usersLoggedOn = @()

	# Sort Descending by lastLogOnDateTime property to get last logon first (topmost)
	# Get more information from first (latest) login user and add information to ToolTip
	$ProcessingLatestCheckinUser = $True
	foreach($LoggedOnUser in ($Script:IntuneManagedDevice.usersLoggedOn | Sort-Object -Property lastLogOnDateTime -Descending)) {
        # Check we have valid GUID
        if([System.Guid]::Parse($LoggedOnUser.userId)) {
			if(($ProcessingLatestCheckinUser -and ($SelectedUser -eq $false)) -or (($ProcessingLatestCheckinUser) -and ($LoggedOnUser.userId -eq $SelectedUser.id))) {
				
				if($LoggedOnUser.userId -eq $Script:PrimaryUser.id) {
					# LoggedOnUser is PrimaryUser or SelectedUser is PrimaryUser
					Write-Verbose "Selected LoggedOnUser is PrimaryUser"
					$Script:LatestCheckedinUser = $Script:PrimaryUser
					
				} else {
					Write-Verbose "Selected LoggedOnUser is NOT PrimaryUser"
					# Get user information with all properties
					$url = "https://graph.microsoft.com/beta/users/$($LoggedOnUser.userId)?`$select=*"
					
					# Add to script wide variable so our Menu action can read userId
					$Script:LatestCheckedinUser = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
				}
				
				# Make sure we don't get here on next round (checked-in user)
				$ProcessingLatestCheckinUser = $False
				
				# This is used to fill information to Recent check-ins textBox in the end of foreach loop
				$AADuser = $Script:LatestCheckedinUser | Select-Object -Property id,userPrincipalName,lastLogOnDateTime
				
				if(-not $SelectedUser) {
					Write-Verbose "Latest signed-in user is $($AADuser.userPrincipalName)"
				}
				
				# Add info to Latest Checked-in User Tooltip
				$WPFtextBlock_LatestCheckInUser_textBox_ToolTip_UPN.Text = $Script:LatestCheckedinUser.userPrincipalName
			
				$BasicInfo = ($Script:LatestCheckedinUser | Select-Object -Property accountEnabled,displayName,userPrincipalName,email,userType,mobilePhone,jobTitle,department,companyName,employeeId,employeeType,streetAddress,postalCode,state,country,officeLocation,usageLocation | Format-List | Out-String).Trim()
				$WPFtextBlock_LatestCheckInUser_textBox_ToolTip_BasicInfo.Text = $BasicInfo

				$proxyAddressesToolTip = ($Script:LatestCheckedinUser.proxyAddresses | Format-List | Out-String).Trim()
				$WPFtextBlock_LatestCheckInUser_textBox_ToolTip_proxyAddresses.Text = $proxyAddressesToolTip

				$otherMailsToolTip = ($Script:LatestCheckedinUser.otherMails | Format-List | Out-String).Trim()
				$WPFtextBlock_LatestCheckInUser_textBox_ToolTip_otherMails.Text = $otherMailsToolTip

				$onPremisesAttributesToolTip = ($Script:LatestCheckedinUser | Select-Object -Property onPremisesSamAccountName, onPremisesUserPrincipalName, onPremisesSyncEnabled, onPremisesLastSyncDateTime, onPremisesDomainName, onPremisesDistinguishedName,onPremisesImmutableId | Format-List | Out-String).Trim()
				$WPFtextBlock_LatestCheckInUser_textBox_ToolTip_onPremisesAttributes.Text = $onPremisesAttributesToolTip

				$onPremisesExtensionAttributesToolTip = ($Script:LatestCheckedinUser.onPremisesExtensionAttributes | Format-List | Out-String).Trim()
				$WPFtextBlock_LatestCheckInUser_textBox_ToolTip_onPremisesExtensionAttributes.Text = $onPremisesExtensionAttributesToolTip

				# Add info to right side Overview tabItem
				$WPFTextBlock_Overview_LatestCheckedInUserName.Text = "Latest or Selected Checked-in User: $($Script:LatestCheckedinUser.userPrincipalName)"
				$WPFTextBox_Overview_LatestCheckedInUser.Text = $BasicInfo
				
				# Add info to right side Latest Checked-In User JSON tabItem textBox
				$WPFLatestCheckedInUser_json_textBox.Text = $Script:LatestCheckedinUser | ConvertTo-Json -Depth 5
				
				# Enable Latest Checked-In User right click menus
				$WPFLatest_CheckIn_User_Menu_Copy.isEnabled = $True
				$WPFLatest_CheckIn_User_Menu_Copy_Menu_OpenLatestCheckInUserInBrowser.isEnabled = $True
				

				# Get Latest LoggedOn User Groups memberOf
				$url = "https://graph.microsoft.com/beta/users/$($Script:LatestCheckedinUser.id)/memberOf?_=1577625591876"
				$Script:LatestCheckedInUserGroupsMemberOf = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
				if($Script:LatestCheckedInUserGroupsMemberOf) {

					# DEBUG $Script:LatestCheckedInUserGroupsMemberOf -> Paste json data to text editor
					#$Script:LatestCheckedInUserGroupsMemberOf | ConvertTo-Json -Depth 4 | Set-ClipBoard

					$Script:LatestCheckedInUserGroupsMemberOf = Add-AzureADGroupGroupTypeExtraProperties $Script:LatestCheckedInUserGroupsMemberOf
					
					Write-Verbose "Add latest checkedin user's AzureAD groups devices and users member count custom properties"
					$Script:LatestCheckedInUserGroupsMemberOf = Add-AzureADGroupDevicesAndUserMemberCountExtraProperties $Script:LatestCheckedInUserGroupsMemberOf

					[Array]$script:LatestCheckedInUserGroupMembershipsObservableCollection = [Array]$Script:LatestCheckedInUserGroupsMemberOf | Sort-Object -Property displayName
					
					$WPFlistView_LatestCheckedInUser_GroupMemberships.Itemssource = $script:LatestCheckedInUserGroupMembershipsObservableCollection

					# Change TabItem Header text to include checked-in user UPN also
					$WPFTabItem_LatestCheckedInUser_GroupMembershipsTAB_Header.Text = "Checked-In User Group Memberships ($($AADuser.userPrincipalName))"

					# Enable LatestCheckedInUser Group Memberships right click menus
					$WPFListView_GridTabItem_LatestCheckedInUser_GroupMembershipsTAB_Menu_Copy_DynamicRules.isEnabled = $True
					$WPFListView_GridTabItem_LatestCheckedInUser_GroupMembershipsTAB_Menu_Copy_JSON.isEnabled = $True
					$WPFListView_GridTabItem_LatestCheckedInUser_GroupMembershipsTAB_Menu_Open_Group_In_Browser.isEnabled = $True

					# Set ToolTip to TabItem Header showing all Azure AD Groups
					$LatestCheckedInUserGroupsMemberOfToolTip = [array]$null
					$Script:LatestCheckedInUserGroupsMemberOf | Sort-Object -Property displayName | Foreach { $LatestCheckedInUserGroupsMemberOfToolTip += "$($_.displayName)`n" }
					$WPFTextBlock_TabItem_LatestCheckedInUser_GroupMembershipsTAB_Header_ToolTip.Text = $LatestCheckedInUserGroupsMemberOfToolTip

				} else {
					Write-Host "Did not find any groups for user $($Script:LatestCheckedinUser.userPrincipalName)"
				}

			} else {
				# Get user information
				$url = "https://graph.microsoft.com/beta/users/$($LoggedOnUser.userId)?`$select=id,displayName,mail,userPrincipalName"
				$AADuser = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
			}
        }
        #$usersLoggedOn += "$($AADuser.mail)`n"
		$usersLoggedOn += "$($AADuser.userPrincipalName)`n"
		$lastLogOnDateTimeLocalTimeZone = (Get-Date $LoggedOnUser.lastLogOnDateTime).ToLocalTime()
		$lastLogOnDateTime = Get-Date $lastLogOnDateTimeLocalTimeZone -Format 'yyyy-MM-dd HH:mm:ss'
        $usersLoggedOn += "$lastLogOnDateTime`n`n"
		
		# DEBUG
		#$script:ComboItemsInApplicationSelectUserComboBoxSource | ConvertTo-Json | Set-Clipboard

		# Add user to User select dropdown selection only if it does not already exists in list
		if($script:ComboItemsInApplicationSelectUserComboBoxSource | Where-Object UserPrincipalName -Like "$($AADuser.UserPrincipalName)*" ) {
			# User is already in list so skip this user
			# so doing nothing
		} else {
			# Add user object to Application Assignments user select dropdown combobox source
			$script:ComboItemsInApplicationSelectUserComboBoxSource += $AADuser
		}
    }
    $WPFIntuneDeviceDetails_RecentCheckins_textBox.Text = $usersLoggedOn
	
	
}




function Get-ApplicationsWithAssignments {
	Param(
		[Parameter(Mandatory=$false)]
		[boolean]$ReloadCacheData = $false
	)
	
	Write-Verbose "Getting Intune Apps information"

    # Check if Apps have changed in Intune after last cached file was loaded
    # We try to get Apps changed after last cache file modified date

    $AppsWithAssignmentsFilePath = "$PSScriptRoot\cache\$TenantId\AllApplicationsWithAssignments.json"

    # Check if AllApplicationsWithAssignments.json file exists
    if((Test-Path "$AppsWithAssignmentsFilePath") -and ($ReloadCacheData -eq $false)) {

        $FileDetails = Get-ChildItem "$AppsWithAssignmentsFilePath"

        # Get timestamp for file AllApplicationsWithAssignments.json
        # We use uformat because Culture can otherwise change separator for time format (change : -> .)
        # Get-Date -uformat %G-%m-%dT%H:%M:%S.0000000Z
        $AppsFileLastWriteTimeUtc = Get-Date $FileDetails.LastWriteTimeUtc -uformat %G-%m-%dT%H:%M:%S.000Z

		# Check how old cache file is
		# We will update all data every Nth days whether there is cache file or not
		$CacheFileOld = New-Timespan (Get-Date $AppsFileLastWriteTimeUtc) (Get-Date)
		$CacheFileOldDays = $CacheFileOld.Days
		
		Write-Verbose "Cache file $AppsWithAssignmentsFilePath is $($CacheFileOld.TotalDays) days old"

		if($CacheFileOldDays -lt $Script:ReloadCacheEveryNDays) {
		
			Write-Verbose "Found cached Apps information (AllApplicationsWithAssignments.json)."
			Write-Verbose "Checking if there are Intune Apps modified after cache file timestamp ($AppsFileLastWriteTimeUtc)"

			# Get MobileApps modified after cache file datetime
			#https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$filter=lastModifiedDateTime%20gt%202019-12-31T00:00:00.000Z&$expand=assignments
			$url = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=lastModifiedDateTime%20gt%20$AppsFileLastWriteTimeUtc&`$expand=assignments&`$top=1000"

			$AllAppsChangedAfterCacheFileDate = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

			if ($AllAppsChangedAfterCacheFileDate) {
				# We found new/changed Apps which we don't have in our cache file so we need to download Apps

				# Future TODO: get changed Apps and migrate that to existing cache file
				# and Always force update Apps after 7 days old cache file

				# For now we don't actually do anything here because next phase will download All Apps and update cache file
				
				# Get new Apps count.
				# This is kind of "stupid" workaround to get count even if there is only 1 object
				# because that is not array and 1 object does not have .Count -property
				$NewAppsCount = $AllAppsChangedAfterCacheFileDate | Measure-Object | Select-Object -ExpandProperty Count
				Write-Verbose "Found $NewAppsCount new or changed Apps from Intune"
				
			} else {
				# We found no changed Apps so our cache file is still valid
				# We can use cached file
				
				Write-Verbose "No new or changed Apps found so we are using cached information"
				
				$AppsWithAssignments = Get-Content "$AppsWithAssignmentsFilePath" | ConvertFrom-Json
				return $AppsWithAssignments
			}
		} else {
			Write-Verbose "Download all data because cache is over 1 days old"
		}
    } 

	Write-Verbose "Download App information from Intune"

	# If we end up here then file either does not exist or we need to update existing cached file

	$url = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$expand=assignments&_=1577625591870"

	$AllAppsWithAssignments = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

	if($AllAppsWithAssignments) {
		# Save to local cache
		# Use -Depth 4 to get all Assignment information also !!!
		$AllAppsWithAssignments | ConvertTo-Json -Depth 4 | Out-File "$AppsWithAssignmentsFilePath" -Force
		
		# Load Application information from cached file always
		$AppsWithAssignments = Get-Content "$AppsWithAssignmentsFilePath" | ConvertFrom-Json
		
		Write-Verbose "Downloaded $($AppsWithAssignments.Count) Apps information from Intune"
		
		return $AppsWithAssignments
	} else {
		Write-Verbose "Did not find any Apps from Intune!"
		
		return $false
	}
}


function Get-MobileAppIntentsForSpecifiedUser {
	Param(
		[Parameter(Mandatory=$true,
			HelpMessage = 'Enter User id')]
			$UserId,
		[Parameter(Mandatory=$true,
			HelpMessage = 'Enter Intune Device id')]
			$IntuneDeviceId
	)

	$script:AppsAssignmentsObservableCollection = @()

	# Get all applications targeted to specific user AND device
	# if there is no Primary User then we get only device targeted applications
	# We will get all device AND user targeted apps. We will need to figure out which apps came from which AzureAD Group targeting

	# Intune original request
	#$url = "https://graph.microsoft.com/beta/users('$($Script:IntuneManagedDevice.userId)')/mobileAppIntentAndStates('$IntuneDeviceId')"

	# Using Primary User id
	$url = "https://graph.microsoft.com/beta/users('$($UserId)')/mobileAppIntentAndStates('$IntuneDeviceId')"

	# Send MSGraph request
	$mobileAppIntentAndStates = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
	if($mobileAppIntentAndStates.mobileAppList) {
		$mobileAppIntentAndStatesMobileAppList = $mobileAppIntentAndStates.mobileAppList
		Write-Verbose "Found $($mobileAppIntentAndStatesMobileAppList.Count) App Intents and States"
	} else {
		$mobileAppIntentAndStatesMobileAppList = $null
		Write-Verbose "Did not find any App Intents and States!"
		return $False
	}

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
		$YodamiittiCustomGroupMembers = 'N/A'
		$AppHadAssignments = $false
		$displayName = $null
		$properties = $null
		$displayNameToolTip = $null

		# Get Application information with Assignment details
		# Get it once and use it many times
		$App = $Script:AppsWithAssignments | Where-Object { $_.id -eq "$($mobileAppIntentAndStatesMobileApp.applicationId)" }
		#$App

		$displayNameToolTip = $App.description

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
			$contextToolTip = $null
			$assignmentGroup = $null
			$YodamiittiCustomGroupMembers = 'N/A'
			$assignmentGroupId = $null
			$AssignmentGroupToolTip = $null

			$assignmentFilterDisplayName = $null
			$assignmentFilterId = $null
			$FilterToolTip = $null
			$FilterMode = $null
			
			# Cast as string so our column sorting works
			# DID NOT WORK for fixing sorting
			$YodamiittiCustomMembershipType = [String]''

			if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
				# Special case for All Users
				$assignmentGroup = 'All Users'
				$context = 'User'
				$contextToolTip = 'Built-in All Users group'
				$AssignmentGroupToolTip = 'Built-in All Users group'

				$YodamiittiCustomGroupMembers = ''

				$IncludeApplicationAssignmentInSummary = $true
			}
			
			if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
				# Special case for All Devices
				$assignmentGroup = 'All Devices'
				$context = 'Device'
				$contextToolTip = 'Built-in All Devices group'
				$AssignmentGroupToolTip = 'Built-in All Devices group'

				$YodamiittiCustomGroupMembers = ''

				$IncludeApplicationAssignmentInSummary = $true
			}
			
			if (($Assignment.target.'@odata.type' -ne '#microsoft.graph.allLicensedUsersAssignmentTarget') -and ($Assignment.target.'@odata.type' -ne '#microsoft.graph.allDevicesAssignmentTarget')) {

				# Group based assignment. We need to get AzureAD Group Name
				# #microsoft.graph.groupAssignmentTarget

				# Test if device is member of this group
				if($Script:deviceGroupMemberships | Where-Object { $_.id -eq $Assignment.target.groupId}) {
					$context = 'Device'
					$contextToolTip = "$($Script:IntuneManagedDevice.deviceName)"

					#$assignmentGroup = $Script:deviceGroupMemberships | Where-Object { $_.id -eq $Assignment.target.groupId} | Select-object -ExpandProperty displayName

					$assignmentGroupObject = $Script:deviceGroupMemberships | Where-Object { $_.id -eq $Assignment.target.groupId}
					
					$assignmentGroup = $assignmentGroupObject.displayName
					$assignmentGroupId = $assignmentGroupObject.id
					
					# Create Group Members column information
					$DevicesCount = $assignmentGroupObject.YodamiittiCustomGroupMembersCountDevices
					$UsersCount = $assignmentGroupObject.YodamiittiCustomGroupMembersCountUsers
					#$YodamiittiCustomGroupMembers = "$DevicesCount devices, $UsersCount users"
					$YodamiittiCustomGroupMembers = ''
					if($DevicesCount -gt 0) { $YodamiittiCustomGroupMembers += "$DevicesCount devices " }
					if($UsersCount -gt 0) { $YodamiittiCustomGroupMembers += "$UsersCount users " }
					
					
					#$GroupType = Add-AzureADGroupGroupTypeExtraProperties $assignmentGroupObject
					$AssignmentGroupToolTip = "$($assignmentGroupObject.membershipRule)"
					
					$YodamiittiCustomMembershipType = $assignmentGroupObject.YodamiittiCustomMembershipType
					
					$IncludeApplicationAssignmentInSummary = $true

				} else {
					# Group not found on member of devicegroups
				}


				# Test if primary user is member of assignment group
				#if($Script:PrimaryUserGroupsMemberOf | Where-Object { $_.id -eq $Assignment.target.groupId}) {
				if(($Script:PrimaryUserGroupsMemberOf | Where-Object { $_.id -eq $Assignment.target.groupId}) -and ($UserId -eq $Script:PrimaryUser.id)) {
					# Process PrimaryUser's group memberships
					
					Write-Verbose "Checking Application assignment group from PrimaryUser's group memberships"
					
					if($assignmentGroup) {
						# Device also is member of this group. Now we got mixed User and Device memberships
						# Maybe not good practise but it is possible

						$context = '_Device/User'
					} else {
						$context = 'User'
						#$assignmentGroup = $Script:PrimaryUserGroupsMemberOf | Where-Object { $_.id -eq $Assignment.target.groupId} | Select-object -ExpandProperty displayName
						
						$assignmentGroupObject = $Script:PrimaryUserGroupsMemberOf | Where-Object { $_.id -eq $Assignment.target.groupId}
						
						$assignmentGroup = $assignmentGroupObject.displayName
						$assignmentGroupId = $assignmentGroupObject.id
						
						# Create Group Members column information
						$DevicesCount = $assignmentGroupObject.YodamiittiCustomGroupMembersCountDevices
						$UsersCount = $assignmentGroupObject.YodamiittiCustomGroupMembersCountUsers
						#$YodamiittiCustomGroupMembers = "$DevicesCount devices, $UsersCount users"
						$YodamiittiCustomGroupMembers = ''
						if($DevicesCount -gt 0) { $YodamiittiCustomGroupMembers += "$DevicesCount devices " }
						if($UsersCount -gt 0) { $YodamiittiCustomGroupMembers += "$UsersCount users " }							

						$AssignmentGroupToolTip = "$($assignmentGroupObject.membershipRule)"
						
						$YodamiittiCustomMembershipType = $assignmentGroupObject.YodamiittiCustomMembershipType
						
					}
					$IncludeApplicationAssignmentInSummary = $true
					$contextToolTip = $Script:PrimaryUser.UserPrincipalName
				} else {
					# Group not found on member of PrimaryUser's groups
				}
				
				
				if(($Script:LatestCheckedInUserGroupsMemberOf | Where-Object { $_.id -eq $Assignment.target.groupId}) -and ($UserId -ne $Script:PrimaryUser.id)) {
					# Process selected User's or recently loggedin user's group memberships

					Write-Verbose "Checking Application assignment group from select user's ($UserId) group memberships"

					# DEBUG
					#Write-Verbose "DEBUG: Process selected User's or recently loggedin user's group memberships"
					#Write-Verbose "DEBUG: UserId: $UserId"
					#Write-Verbose "DEBUG: Assignment.target.groupId: $($Assignment.target.groupId)"
					#$Script:LatestCheckedInUserGroupsMemberOf | ConvertTo-Json -Depth 5 | Set-Clipboard
					
					if($assignmentGroup) {
						# Device also is member of this group. Now we got mixed User and Device memberships
						# Maybe not good practise but it is possible

						$context = '_Device/User'
					} else {
						$context = 'User'
						#$assignmentGroup = $Script:PrimaryUserGroupsMemberOf | Where-Object { $_.id -eq $Assignment.target.groupId} | Select-object -ExpandProperty displayName
						
						$assignmentGroupObject = $Script:LatestCheckedInUserGroupsMemberOf | Where-Object { $_.id -eq $Assignment.target.groupId}
						
						$assignmentGroup = $assignmentGroupObject.displayName
						$assignmentGroupId = $assignmentGroupObject.id
						
						# Create Group Members column information
						$DevicesCount = $assignmentGroupObject.YodamiittiCustomGroupMembersCountDevices
						$UsersCount = $assignmentGroupObject.YodamiittiCustomGroupMembersCountUsers
						#$YodamiittiCustomGroupMembers = "$DevicesCount devices, $UsersCount users"
						$YodamiittiCustomGroupMembers = ''
						if($DevicesCount -gt 0) { $YodamiittiCustomGroupMembers += "$DevicesCount devices " }
						if($UsersCount -gt 0) { $YodamiittiCustomGroupMembers += "$UsersCount users " }							

						$AssignmentGroupToolTip = "$($assignmentGroupObject.membershipRule)"
						
						$YodamiittiCustomMembershipType = $assignmentGroupObject.YodamiittiCustomMembershipType
						
					}
					$IncludeApplicationAssignmentInSummary = $true
					$contextToolTip = $Script:LatestCheckedinUser.UserPrincipalName
				} else {
					# Group not found on member of selected or latest signed-in user
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
				#$assignmentFilterDisplayName = $AllIntuneFilters | Where-Object { $_.id -eq $assignmentFilterId } | Select-Object -ExpandProperty displayName
				
				$assignmentFilterObject = $AllIntuneFilters | Where-Object { $_.id -eq $assignmentFilterId }

				$assignmentFilterDisplayName = $assignmentFilterObject.displayName
				$assignmentFilterId = $assignmentFilterObject.id
				$FilterToolTip = $assignmentFilterObject.rule

				$FilterMode = $Assignment.target.deviceAndAppManagementAssignmentFilterType
				if($FilterMode -eq 'None') {
					$FilterMode = $null
				}

				# Cast variable types to make sure column click based sorting works
				# Sorting may break if there are different kind of objects
				$properties = @{
					context                          = [String]$context
					contextToolTip					 = [String]$contextToolTip
					odatatype                        = [String]$odatatype
					displayname                      = [String]$displayName
					version                          = [String]$mobileAppIntentAndStatesMobileApp.displayVersion
					assignmentIntent                 = [String]$assignmentIntent
					IncludeExclude                   = [String]$AppIncludeExclude
					assignmentGroup                  = [String]$assignmentGroup
					YodamiittiCustomGroupMembers     = [String]$YodamiittiCustomGroupMembers
					assignmentGroupId                = [String]$assignmentGroupId
					installState                     = [String]$mobileAppIntentAndStatesMobileApp.installState
					lastModifiedDateTime             = $App.lastModifiedDateTime
					YodamiittiCustomMembershipType   = [String]$YodamiittiCustomMembershipType
					id                               = $App.id
					filter							 = [String]$assignmentFilterDisplayName
					filterId						 = [String]$assignmentFilterId
					filterMode						 = [String]$FilterMode
					filterTooltip                    = [String]$FilterTooltip
					AssignmentGroupToolTip 			 = [String]$AssignmentGroupToolTip
					displayNameToolTip               = [String]$displayNameToolTip
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
			
			# Set variable which we return from this function
			$UnknownAssignmentGroupFound = $true
			
			# One option is that our cache data was not updated even if data was updated in Intune
			# Some configuration policies may not change lastModifiedDateTime property if only assignments were changed

			$context = '_unknown'
 
			# App Intent requiredInstall is different than App Assignment so we remove word Install
			$assignmentIntent = $mobileAppIntentAndStatesMobileApp.mobileAppIntent
			$assignmentIntent = $assignmentIntent.Replace('Install','')

			$AppIncludeExclude = ''
			$assignmentGroup = 'unknown (possible nested group or removed assignment)'
			$YodamiittiCustomGroupMembers = 'N/A'
	
			# Cast variable types to make sure column click based sorting works
			# Sorting may break if there are different kind of objects
			$properties = @{
				context                          = [String]$context
				contextToolTip					 = [String]''
				odatatype                        = [String]$odatatype
				displayname                      = [String]$displayName
				version                          = [String]$mobileAppIntentAndStatesMobileApp.displayVersion
				assignmentIntent                 = [String]$assignmentIntent
				IncludeExclude                   = [String]$AppIncludeExclude
				assignmentGroup                  = [String]$assignmentGroup
				YodamiittiCustomGroupMembers     = [String]$YodamiittiCustomGroupMembers
				assignmentGroupId                = [String]$null
				installState                     = [String]$mobileAppIntentAndStatesMobileApp.installState
				lastModifiedDateTime             = $App.lastModifiedDateTime
				YodamiittiCustomMembershipType   = [String]''
				id                               = $App.id
				filter							 = [String]''
				filterId						 = [String]$null
				filterMode						 = [String]''
				filterTooltip                    = [String]''
				AssignmentGroupToolTip 			 = [String]''
				displayNameToolTip               = [String]''
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
		# Cast as array because otherwise it will fail
		$WPFlistView_ApplicationAssignments.Itemssource = [array]$script:AppsAssignmentsObservableCollection
	}

	# If we got here then we should have at least 1 App Intent and State
	return $mobileAppIntentAndStatesMobileAppList.Count
}


function Download-IntunePostTypeReport {
	Param(
		[Parameter(Mandatory=$true,
			HelpMessage = 'Enter Graph API Url')]
		$GraphAPIUrl,
		[Parameter(Mandatory=$true,
			HelpMessage = 'Enter Graph API Post request')]
		$GraphAPIPostRequest
	)

	# Initialize variables
	# Not actually needed here
	# but helps coder to think about loop logic
	$MSGraphRequest = $null
	$ConfigurationPoliciesReportForDevice = @()
	$count = $null

	do {
		Start-Sleep -Milliseconds $TimeOutBetweenGraphAPIRequests

		$GraphAPIPostRequestJSON = $GraphAPIPostRequest | ConvertFrom-Json

		$top = $GraphAPIPostRequestJSON.top
		$skip = $GraphAPIPostRequestJSON.skip

		# DEBUG
		#Write-Verbose "`$top=$top"
		#Write-Verbose "`$skip=$skip"

		$MSGraphRequest = Invoke-MSGraphRequest -Url $url -Content $GraphAPIPostRequest.ToString() -HttpMethod 'POST'
		$Success = $?

		if($Success) {
			#Write-Verbose "Success"

			# Objectify report results
			$MSGraphRequestObjectified = Objectify_JSON_Schema_and_Data_To_PowershellObjects $MSGraphRequest
			
			# Save results to variable
			$ConfigurationPoliciesReportForDevice += $MSGraphRequestObjectified

			# Get Count of results
			$count = $MSGraphRequestObjectified.Count

			if($count -ge $top) {
				# Increase report skip-value with amount of results we got earlier (should be same as top)
				# to get next batch of results
				$skip += $count

				# Increase count in json and convert to text
				#$GraphAPIPostRequestJSON.top = $top
				$GraphAPIPostRequestJSON.skip = $skip

				# Convert json to text
				$GraphAPIPostRequest = $GraphAPIPostRequestJSON | ConvertTo-Json -Depth 3

			} else {
				# Got all results

				#Write-Verbose "DEBUG"
				#Write-Verbose "Got all policy report results"
				#Write-Verbose "`$count=$count"
				#Write-Verbose "`$top=$top"
				#Write-Verbose "`$skip=$skip"
				Write-Verbose "Found $($ConfigurationPoliciesReportForDevice.Count) assignment objects"
			}

		} else {
			# Invoke-MSGraphRequest failed
			Write-Error "Error getting Intune device Configuration Assignment information"
			return 1
		}
	} while ($count -ge $top)

	return $ConfigurationPoliciesReportForDevice
}


function Download-IntuneConfigurationProfiles2 {
	Param(
		[Parameter(Mandatory=$true,
			HelpMessage = 'Enter Graph API Url')]
		$GraphAPIUrl,
		[Parameter(Mandatory=$true,
			HelpMessage = 'Enter JSON Cache File Name')]
		$jsonCacheFileName,
		[Parameter(Mandatory=$false)]
		[boolean]$ReloadCacheData = $false
	)

	$jsonCacheFilePath = "$PSScriptRoot\cache\$TenantId\$jsonCacheFileName"
	
	# Intune Policies to download including Assignment information
	# Note that Endpint Security configurations (intents) do not expand assignments!!!

	Write-Verbose ''
	Write-Verbose "Getting Intune configuration for url: $GraphAPIUrl"

    # Check if Configuration have changed in Intune after last cached file was loaded
    # We try to get Configurations changed after last cache file modified date

    # Check if json cache file exists
	# And continue only if $ReloadCacheData=$False
    if((Test-Path "$jsonCacheFilePath") -and ($ReloadCacheData -eq $false)) {

        $FileDetails = Get-ChildItem "$jsonCacheFilePath"

        # Get timestamp for file cache file
        # We use uformat because Culture can otherwise change separator for time format (change : -> .)
        # Get-Date -uformat %G-%m-%dT%H:%M:%S.0000000Z
        $CacheFileLastWriteTimeUtc = Get-Date $FileDetails.LastWriteTimeUtc -uformat %G-%m-%dT%H:%M:%S.000Z

		# Check how old cache file is
		# We will update all data every Nth days whether there is cache file or not
		$CacheFileOld = New-Timespan (Get-Date $CacheFileLastWriteTimeUtc) (Get-Date)
		$CacheFileOldDays = $CacheFileOld.Days

		Write-Verbose "Cache file $jsonCacheFilePath is $($CacheFileOld.TotalDays) days old"

		if($CacheFileOldDays -lt $Script:ReloadCacheEveryNDays) {
			Write-Verbose "Found cached Configuration file ($jsonCacheFilePath)."
			Write-Verbose "Checking if there are Intune Configurations modified after cache file timestamp ($CacheFileLastWriteTimeUtc)"

			# Get Configurations modified after cache file datetime
			$url = "$GraphAPIUrl&`$filter=lastModifiedDateTime%20gt%20$CacheFileLastWriteTimeUtc"
			
			$AllConfigurationsChangedAfterCacheFileDate = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

			if ($AllConfigurationsChangedAfterCacheFileDate) {
				# We found new/changed Configurations which we don't have in our cache file so we need to download Configs

				# Future TODO: get changed data and migrate that to existing cache file
				# and Always force update data after 7 days old cache file

				# For now we don't actually do anything here because next phase will download All data and update cache file
				
				# Get new data count.
				# This is kind of "stupid" workaround to get count even if there is only 1 object
				# because that is not array and 1 object does not have .Count -property
				$NewConfigurationsCount = $AllConfigurationsChangedAfterCacheFileDate | Measure-Object | Select-Object -ExpandProperty Count
				Write-Verbose "Found $NewConfigurationsCount new Configurations from Intune"
				
			} else {
				# We found no changed Configurations so our cache file is still valid
				# We can use cached file
				
				Write-Verbose "No new Configurations found so we are using cached information"
				
				[array]$ConfigurationsWithAssignments = Get-Content "$jsonCacheFilePath" | ConvertFrom-Json
				return $ConfigurationsWithAssignments
			}
		} else {
			Write-Verbose "Download all data because cache is over 1 days old"
		}
    }

	Write-Verbose "Download Configuration information from Intune"

	# If we end up here then file either does not exist or we need to update existing cached file
	# or $ReloadCacheData was true
	[array]$ConfigurationsWithAssignments = Invoke-MSGraphGetRequestWithMSGraphAllPages $GraphAPIUrl

	if($ConfigurationsWithAssignments) {
		# Save to local cache
		# Use -Depth 6 to get all Assignment information also !!!
		[array]$ConfigurationsWithAssignments | ConvertTo-Json -Depth 6 | Out-File "$jsonCacheFilePath" -Force
		
		# Load Configuration information from cached file always
		[array]$ConfigurationsWithAssignments = Get-Content "$jsonCacheFilePath" | ConvertFrom-Json
		
		Write-Verbose "Downloaded $($ConfigurationsWithAssignments.Count) Configuration information from Intune"
		
		return $ConfigurationsWithAssignments
	} else {
		Write-Verbose "Did not find any Configurations from Intune!"
		
		return $false
	}
}


function Download-IntuneFilters {
	try {
		$url = 'https://graph.microsoft.com/beta/deviceManagement/assignmentFilters?$select=*'
		$MSGraphRequest = Invoke-MSGraphRequest -Url $url -HttpMethod 'GET'
		$Success = $?

		if (-not ($Success)) {
			Write-Error "Error downloading Intune filters information"
			return $null
		} else {
			$AllIntuneFilters = Get-MSGraphAllPages -SearchResult $MSGraphRequest	
		}

		Write-Verbose "Found $($AllIntuneFilters.Count) Intune filters"

		if($AllIntuneFilters) {
			# Save to local cache
			# Use -Depth 6 to get all information (default 3 should be enough in this case)
			$jsonCacheFilePath = "$PSScriptRoot\cache\$TenantId\filters.json"
			$AllIntuneFilters | ConvertTo-Json -Depth 6 | Out-File "$jsonCacheFilePath" -Force
			
			# Load Configuration information from cached file always
			$AllIntuneFilters = Get-Content "$jsonCacheFilePath" | ConvertFrom-Json
			
			return $AllIntuneFilters
		} else {
			Write-Verbose "Did not find any Intune Filters!"
			
			return $null
		}

    } catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download Intune filters information"
		return $null
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


function Add-AzureADGroupGroupTypeExtraProperties {
	Param(
		[Parameter(Mandatory=$true,
			ValueFromPipeline=$true,
			ValueFromPipelineByPropertyName=$true, 
			Position=0)]
		$AzureADGroups
	)

	# DEBUG Export $group to clipboard for testing off the script
	#$AzureADGroups | ConvertTo-Json -Depth 5 | Set-Clipboard

	# Add new properties groupType and MembershipType
	foreach($group in $AzureADGroups) {

		$GroupType = 'unknown'
		if($group.groupTypes -contains 'Unified') {
			# Group is Office365 group
			$group | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupType -Value 'Office365'
		} else {
			# Group is either security group or distribution group

			if($group.securityEnabled -and (-not $group.mailEnabled)) {
				# Group is security group
				$group | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupType -Value 'Security'
			}
			
			if((-not $group.securityEnabled) -and $group.mailEnabled) {
				# Group is Distribution group
				$group | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupType -Value 'Distribution'
			}
		}


		# Check if group is directoryRole which is not actual AzureAD Group
		if($group.'@odata.type' -eq '#microsoft.graph.directoryRole') {
			# Group is NOT security group at all
			# DirectoryRoles are not Azure AD groups
			$group | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupType -Value 'DirectoryRole'
		}


		if($group.groupTypes -contains 'DynamicMembership') {
			# Dynamic group
			$group | Add-Member -MemberType noteProperty -Name YodamiittiCustomMembershipType -Value 'Dynamic'
		} else {
			# Static group
			$group | Add-Member -MemberType noteProperty -Name YodamiittiCustomMembershipType -Value 'Static'
		}
	}

	return $AzureADGroups
}


function Add-AzureADGroupDevicesAndUserMemberCountExtraProperties {
	Param(
		[Parameter(Mandatory=$true,
			ValueFromPipeline=$true,
			ValueFromPipelineByPropertyName=$true, 
			Position=0)]
		$AzureADGroups
	)

	Write-Verbose "Getting AzureAD groups membercount for $($AzureADGroups.Count) groups"

	for ($i=0; $i -lt $AzureADGroups.count; $i+=20){

		# Create requests hashtable
		$requests_devices_count = @{}
		$requests_users_count = @{}

		# Create elements array inside hashtable
		$requests_devices_count.requests = @()
		$requests_users_count.requests = @()

		# Create max 20 requests in for-loop
		# For-loop will end automatically when loop counter is same as total count of $AzureADGroups
		for ($a=$i; (($a -lt $i+20) -and ($a -lt $AzureADGroups.count)); $a+=1) {

			if(($AzureADGroups[$a]).'@odata.type' -eq '#microsoft.graph.directoryRole') {
				# Azure DirectoryRole is not AzureAD Group
				$GraphAPIBatchEntry_DevicesCount = @{
					id = ($a+1).ToString()
					"method" = "GET"
					"url" = "/directoryRoles/$(($AzureADGroups[$a]).id)"
				}

			} else {
				# We should have AzureAD Group
				$GraphAPIBatchEntry_DevicesCount = @{
					id = ($a+1).ToString()
					"method" = "GET"
					"url" = "/groups/$(($AzureADGroups[$a]).id)/transitivemembers/microsoft.graph.device/`$count?ConsistencyLevel=eventual"
				}
			}

			# Add GraphAPI Batch entry to requests array
			$requests_devices_count.requests += $GraphAPIBatchEntry_DevicesCount

			if(($AzureADGroups[$a]).'@odata.type' -eq '#microsoft.graph.directoryRole') {
				# Azure DirectoryRole is not AzureAD Group
				$GraphAPIBatchEntry_UsersCount = @{
					id = ($a+1).ToString()
					"method" = "GET"
					"url" = "/directoryRoles/$(($AzureADGroups[$a]).id)"
				}
			} else {
				# We should have AzureAD Group
				$GraphAPIBatchEntry_UsersCount = @{
					id = ($a+1).ToString()
					"method" = "GET"
					"url" = "/groups/$(($AzureADGroups[$a]).id)/transitivemembers/microsoft.graph.user/`$count?ConsistencyLevel=eventual"
				}
			}

			
			# Add GraphAPI Batch entry to requests array
			$requests_users_count.requests += $GraphAPIBatchEntry_UsersCount
			
			# DEBUG/double check index numbers and groupNames
			#Write-Host "`$a=$a   `$i=$i    GroupName=$($AzureADGroups[$a].displayName)"
		}

		# DEBUG
		#$requests_devices_count | ConvertTo-Json
		$requests_devices_count_JSON = $requests_devices_count | ConvertTo-Json

		$url = 'https://graph.microsoft.com/beta/$batch'
		$MSGraphRequest = Invoke-MSGraphRequest -Url $url -Content $requests_devices_count_JSON.ToString() -HttpMethod 'POST'
		$Success = $?

		if($Success) {
			#Write-Host "Success"
		} else {
			# Invoke-MSGraphRequest failed
			Write-Error "Error getting AzureAD groups devices count"
			return 1
		}

		# Get AllMSGraph pages
		# This is also workaround to get objects without assigning them from .Value attribute
		$AzureADGroups_Devices_MemberCount_Batch_Result = Get-MSGraphAllPages -SearchResult $MSGraphRequest
		$Success = $?

		if($Success) {
			#Write-Host "Success"
		} else {
			# Invoke-MSGraphRequest failed
			Write-Error "Error getting AzureAD groups devices count"
			return 1
		}
		
		# DEBUG
		#$AzureADGroups_Devices_MemberCount_Batch_Result

		# Process results for devices count batch requests
		Foreach ($response in $AzureADGroups_Devices_MemberCount_Batch_Result.responses) {
			$GroupArrayIndex = $response.id - 1
			if($response.status -eq 200) {
				
				if(($AzureADGroups[$GroupArrayIndex]).'@odata.type' -eq '#microsoft.graph.directoryRole') {
					# DEBUG
					#Write-Verbose "AzureAD directoryRole (arrayIndex=$GroupArrayIndex) $($AzureADGroups[$GroupArrayIndex].displayName)"

					$AzureADGroups[$GroupArrayIndex] | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupMembersCountDevices -Value 'N/A'
				} else {
					# DEBUG
					#Write-Verbose "AzureAD group (arrayIndex=$GroupArrayIndex) $($AzureADGroups[$GroupArrayIndex].displayName) adding devices count property: $($response.body)"
					
					$AzureADGroups[$GroupArrayIndex] | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupMembersCountDevices -Value $response.body					
				}
			} else {
				Write-Error "Error getting devices count for AzureAD group $($AzureADGroups[$GroupArrayIndex].displayName)"
				Write-Error "$($response | ConvertTo-Json)"
			}
		}


		$requests_users_count_JSON = $requests_users_count | ConvertTo-Json

		$url = 'https://graph.microsoft.com/beta/$batch'
		$MSGraphRequest = Invoke-MSGraphRequest -Url $url -Content $requests_users_count_JSON.ToString() -HttpMethod 'POST'
		$Success = $?

		if($Success) {
			#Write-Host "Success"
		} else {
			# Invoke-MSGraphRequest failed
			Write-Error "Error getting AzureAD groups users count"
			return 1
		}

		# Get AllMSGraph pages
		# This is also workaround to get objects without assigning them from .Value attribute
		$AzureADGroups_Users_MemberCount_Batch_Result = Get-MSGraphAllPages -SearchResult $MSGraphRequest
		$Success = $?

		if($Success) {
			#Write-Host "Success"
		} else {
			# Invoke-MSGraphRequest failed
			Write-Error "Error getting AzureAD groups users count"
			return 1
		}
		
		# DEBUG
		#$AzureADGroups_Users_MemberCount_Batch_Result

		# Process results for devices count batch requests
		Foreach ($response in $AzureADGroups_Users_MemberCount_Batch_Result.responses) {
			$GroupArrayIndex = $response.id - 1
			if($response.status -eq 200) {
				
				if(($AzureADGroups[$GroupArrayIndex]).'@odata.type' -eq '#microsoft.graph.directoryRole') {
					# DEBUG
					#Write-Verbose "AzureAD directoryRole (arrayIndex=$GroupArrayIndex) $($AzureADGroups[$GroupArrayIndex].displayName)"

					# Change "AzureAD Group" json to actual real directoryRole json which we just got from batch request
					$AzureADGroups[$GroupArrayIndex] = $response.body
					
					$AzureADGroups[$GroupArrayIndex] | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupMembersCountUsers -Value 'N/A'
					
					# We need to add below properties again because we just replace whole object so we lost earlier customProperties
					$AzureADGroups[$GroupArrayIndex] | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupMembersCountDevices -Value 'N/A' -Force
					$AzureADGroups[$GroupArrayIndex] | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupType -Value 'DirectoryRole' -Force
				} else {
					# DEBUG
					#Write-Verbose "AzureAD group (arrayIndex=$GroupArrayIndex) $($AzureADGroups[$GroupArrayIndex].displayName) adding users count property: $($response.body)"
					
					$AzureADGroups[$GroupArrayIndex] | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupMembersCountUsers -Value $response.body
				}
				
				
			} else {
				Write-Error "Error getting users count for AzureAD group $($AzureADGroups[$GroupArrayIndex].displayName)"
				Write-Error "$($response | ConvertTo-Json)"
			}
		}
		
	}     

	return $AzureADGroups
}

function Fix-UrlSpecialCharacters {
	Param (
		[Parameter(Mandatory=$true,
			ValueFromPipeline=$true,
			ValueFromPipelineByPropertyName=$true, 
			Position=0)]
		$url=$null
	)

	if($url) {
		# Fix url special characters
		$url = $url.Replace(' ', '%20')
		$url = $url.Replace('"', '%22')
		$url = $url.Replace("'", '%27')
		$url = $url.Replace("\", '%5C')
		$url = $url.Replace("@", '%40')
		$url = $url.Replace('ä', '%C3%A4')
		$url = $url.Replace('Ä', '%C3%84')
		$url = $url.Replace('ö', '%C3%B6')
		$url = $url.Replace('Ö', '%C3%96')
		$url = $url.Replace('å', '%C3%A5')
		$url = $url.Replace('Å', '%C3%85')
	}
	
	return $url
}


function Update-QuickFilters {
	
	# Quick Filter search rules
	# You can edit/add/remove Quick Filter rules here. Just don't break the syntax ;)
	$QuickSearchFiltersJSON = @"
[
    {
        "QuickFilterName":  "Search by deviceName, serialNumber, emailAddress, OS or id",
        "QuickFilterGraphAPIFilter":  null
    },
	{
        "QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Devices Synced     in last 15 minutes", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(lastSyncDateTime gt $((Get-Date).ToUniversalTime().AddMinutes(-15) | Get-Date -uformat %G-%m-%dT%H:%M:%S.000Z))&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Devices Synced     in last  1 hour", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(lastSyncDateTime gt $((Get-Date).ToUniversalTime().AddHours(-1) | Get-Date -uformat %G-%m-%dT%H:%M:%S.000Z))&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Devices Synced     in last 24 hours", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(lastSyncDateTime gt $((Get-Date).ToUniversalTime().AddHours(-24) | Get-Date -uformat %G-%m-%dT%H:%M:%S.000Z))&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Devices Synced     today (since midnight)", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(lastSyncDateTime gt $((Get-Date -Hour 0 -Minute 0 -Second 0).ToUniversalTime() | Get-Date -uformat %G-%m-%dT%H:%M:%S.000Z))&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Devices Synced     in last  7 days", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(lastSyncDateTime gt $((Get-Date).AddDays(-7) | Get-Date -Format 'yyyy-MM-dd')T00:00:00.000Z)&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Devices Synced     in last 30 days", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(lastSyncDateTime gt $((Get-Date).AddDays(-30) | Get-Date -Format 'yyyy-MM-dd')T00:00:00.000Z)&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Devices Enrolled   in last 15 minutes", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(enrolleddatetime gt $((Get-Date).ToUniversalTime().AddMinutes(-15) | Get-Date -uformat %G-%m-%dT%H:%M:%S.000Z))&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Devices Enrolled   in last  1 hour", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(enrolleddatetime gt $((Get-Date).ToUniversalTime().AddHours(-1) | Get-Date -uformat %G-%m-%dT%H:%M:%S.000Z))&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Devices Enrolled   today (since midnight)", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(enrolleddatetime gt $((Get-Date -Hour 0 -Minute 0 -Second 0).ToUniversalTime() | Get-Date -uformat %G-%m-%dT%H:%M:%S.000Z))&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Devices Enrolled   in last  7 days", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(enrolleddatetime gt $((Get-Date).AddDays(-7) | Get-Date -Format 'yyyy-MM-dd')T00:00:00.000Z)&`$top=$GraphAPITop"
    },	
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Devices Enrolled   in last 30 days", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(enrolleddatetime gt $((Get-Date).AddDays(-30) | Get-Date -Format 'yyyy-MM-dd')T00:00:00.000Z)&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Windows 10 Devices", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(contains(osVersion,'10.0.1'))&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Windows 11 Devices", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(contains(osVersion,'10.0.2'))&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Windows devices", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "contains(activationlockbypasscode,%20'Windows')&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: Android devices", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "contains(activationlockbypasscode,%20'Android')&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: iPhone devices", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "contains(activationlockbypasscode,%20'iPhone')&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: iPad devices", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "contains(activationlockbypasscode,%20'iPad')&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-55} {1,-20}" -f "Quick filter: macOS devices", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "contains(activationlockbypasscode,%20'macOS')&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-32} {1,-22} {2,-20}" -f "Quick filter: Compliance","Compliant", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(complianceState eq 'compliant')&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-32} {1,-22} {2,-20}" -f "Quick filter: Compliance", "Non-compliant", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(complianceState eq 'noncompliant')&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-32} {1,-22} {2,-20}" -f "Quick filter: Compliance", "Unknown", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(complianceState eq 'unknown')&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-32} {1,-22} {2,-20}" -f "Quick filter: Ownership", "Company devices", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(ownerType eq 'company')&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-32} {1,-22} {2,-20}" -f "Quick filter: Ownership", "Personal devices", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(ownerType eq 'personal')&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-32} {1,-22} {2,-20}" -f "Quick filter: Manufacturer", "Microsoft Corporation", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(manufacturer eq 'Microsoft Corporation')&`$top=$GraphAPITop"
    },
	{
		"QuickFilterName":  "$("{0,-32} {1,-22} {2,-20}" -f "Quick filter: Model", "Virtual Machine", "(Max $GraphAPITop devices)")",
        "QuickFilterGraphAPIFilter":  "(model eq 'Virtual Machine')&`$top=$GraphAPITop"
    }
]
"@

	$QuickSearchFilters = $QuickSearchFiltersJSON | ConvertFrom-Json
	
	return $QuickSearchFilters
}




###################################################################################################################

function Search-IntuneDevices {

	[array]$AllSearchResults = $null
	
	Write-Host "Searching devices"
	
	$WPFlabel_GridIntuneDeviceDetailsBorderTop_FoundXDevices.Visibility = 'Hidden'
	
	# Check if selectedItem is Intune device object type which has id property
	# This device id property can be used as is
	$SearchString = $WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.SelectedItem

	if($SearchString.id) {
		# We got Intune device object because there is property id
		$SearchString = $SearchString.id
	} else {

		# If text was typed manually (not used QuickFilter)
		if($SearchString.QuickFilterGraphAPIFilter -eq $null) {

			# Search string should work as is
			$SearchString = $WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.Text
		
			# Fix special characters so url will work (for example spaces)
			$SearchString = Fix-UrlSpecialCharacters $SearchString
			
			Write-Verbose "Searchstring: $SearchString"
		}
	}


	if(-not $SearchString) {
		$WPFIntuneDeviceDetails_textBox_DeviceName.Text = "Search text was empty"
		$WPFIntuneDeviceDetails_textBox_DeviceName.Foreground = "red"
		return
	}


	# If $SearchString is guid then we will get that specific device
	try {
		[System.Guid]::Parse($SearchString) | Out-Null
			$SearchStringIsValidGUID = $true
		} catch {
			$SearchStringIsValidGUID = $false
		}

	if($SearchStringIsValidGUID) {
		# SearchString is valid guid
		# We get here if source was Intune device object

		Write-Verbose "SearchString ($SearchString) is valid guid"

		# Clear result ComboBox just in case there was previous search results
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.ItemsSource = $null
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.Items.Clear()
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.IsDropDownOpen = $false
		
		# Disable Create report button
		$WPFButton_GridIntuneDeviceDetailsBorderTop_CreateReport.isEnabled = $False
		
		# Set button font to normal
		$WPFButton_GridIntuneDeviceDetailsBorderTop_CreateReport.FontWeight = "Normal"

		[array]$AllSearchResults = Get-IntuneManagedDevice -managedDeviceId $SearchString
		$Success = $?
		if($Success -and $AllSearchResults) {
			Write-Verbose "Found Intune device: $($AllSearchResults.deviceName)"
			foreach($device in $AllSearchResults) {
				$lastSyncDateTime = $device.lastSyncDateTime
				
				$DeviceLastSyncDateTimeDaysAgo = New-TimeSpan (Get-Date) (Get-Date $lastSyncDateTime) | Select-Object -ExpandProperty Days
				$DeviceSyncStringToSearch = "Last sync $DeviceLastSyncDateTimeDaysAgo days ago"
				
				# Add new attribute to object
				$searchStringDeviceProperty = "{0,-20} {1,-8} {2,4} {3,8}" -f $device.deviceName, "Last sync", $DeviceLastSyncDateTimeDaysAgo, "days ago"
			
				$device | Add-Member -MemberType noteProperty -Name searchStringDeviceProperty -Value $searchStringDeviceProperty
			}
			
		} else {
			Write-Verbose "Error finding Intune DeviceId: $SearchString"
			$WPFIntuneDeviceDetails_textBox_DeviceName.Text = "Error finding Intune DeviceId: $SearchString"
			$WPFIntuneDeviceDetails_textBox_DeviceName.Foreground = "red"
		}
		
	} else {
		# Search string is NOT valid guid or Intune device object (which has id property)
		# SearchString is either typed text or Quick Filter from dropdown selection

		# Clear result ComboBox just in case there was previous search results
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.ItemsSource = $null
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.Items.Clear()
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.IsDropDownOpen = $false
		
		# Disable Create report button
		$WPFButton_GridIntuneDeviceDetailsBorderTop_CreateReport.isEnabled = $False
		
		# Set button font to normal
		$WPFButton_GridIntuneDeviceDetailsBorderTop_CreateReport.FontWeight = "Normal"

		if($SearchString.QuickFilterGraphAPIFilter) {
			# Quick Filter was selected from ComboBox Dropdown

			# Update QuickFilters
			$Script:QuickSearchFilters = Update-QuickFilters

			# Get updated datetime value for property QuickFilterGraphAPIFilter
			$GraphAPIFilter = $Script:QuickSearchFilters | Where-Object { $_.QuickFilterName -eq $SearchString.QuickFilterName } | Select-Object -ExpandProperty QuickFilterGraphAPIFilter

			$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=$GraphAPIFilter&`$select=id,deviceName,usersLoggedOn,lastSyncDateTime,operatingSystem,deviceType,enrolledDateTime,lastSyncDateTime,Manufacturer,Model,SerialNumber,userPrincipalName"
			
			# Fix special characters so url will work (for example spaces)
			$url = Fix-UrlSpecialCharacters $url
			
			Write-Verbose "Making search based on quick filter: $url"
			
		} else {	
			Write-Verbose "Making search based on typed text: $SearchString"
			# "General" search

			# More information so we can populate ToolTip information
			$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=contains(activationlockbypasscode,%20'$SearchString')&`$select=id,deviceName,usersLoggedOn,lastSyncDateTime,operatingSystem,deviceType,enrolledDateTime,lastSyncDateTime,Manufacturer,Model,SerialNumber,userPrincipalName&`$Top=$GraphAPITop"
		}

		
		# Cast as array
		[array]$AllSearchResults = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
		
		# Sort by lastSyncDateTime Descending
		$AllSearchResults = $AllSearchResults | Sort-Object -Property lastSyncDateTime -Descending

		foreach($device in $AllSearchResults) {
			$lastSyncDateTime = $device.lastSyncDateTime
			
			$DeviceLastSyncDateTimeDaysAgo = New-TimeSpan (Get-Date) (Get-Date $lastSyncDateTime) | Select-Object -ExpandProperty Days
			$DeviceSyncStringToSearch = "Last sync $DeviceLastSyncDateTimeDaysAgo days ago"
			
			# Add new attribute to object

			# Align output formatting using -f Format Operator
			$searchStringDeviceProperty = "{0,-20} {1,-8} {2,4} {3,8}" -f $device.deviceName, "Last sync", $DeviceLastSyncDateTimeDaysAgo, "days ago"
		
			$device | Add-Member -MemberType noteProperty -Name searchStringDeviceProperty -Value $searchStringDeviceProperty
		}


		# If searchString contains email address then do another search
		if((($SearchString -like "*@*.*") -or ($SearchString -like "*%40*.*")) -and (-not $SearchString.QuickFilterGraphAPIFilter)) {
			# SearchString is email address
			Write-Verbose "SearchString contains email address"

			# Find AzureAD User
			$url = "https://graph.microsoft.com/beta/users?`$filter=userPrincipalName%20eq%20'$SearchString'&`$select=id,mail,userPrincipalName"

			# NOTE! Below url will break next search where we get devices where user has logged on
			# Do not use this as is! But left here because in the future we might also like to search ownedDevices
			# Expand ownedDevices
			# Another expandable collection would be $Expand=registeredDevices
			#$url = "https://graph.microsoft.com/beta/users?`$filter=userPrincipalName%20eq%20'$SearchString'&`$select=id,mail,userPrincipalName&`$Expand=ownedDevices"
			
			Write-Verbose "Get Azure ADUser information with url: $url"

			$AzureADUser = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

			if($AzureADUser.id) {
				if($AzureADUser -is [array]) {
					Write-Verbose "Email search resulted multiple Azure AD Users so can't search devices where users has logged-on!"
				} else {
					# Found 1 user
					
					# More information so we can populate ToolTip information
					$url =	"https://graph.microsoft.com/beta/users/$($AzureADUser.id)/getLoggedOnManagedDevices?`$select=id,deviceName,usersLoggedOn,lastSyncDateTime,operatingSystem,deviceType,enrolledDateTime,lastSyncDateTime,Manufacturer,Model,SerialNumber,userPrincipalName"
					
					Write-Verbose "Get devices where user has logged on with url: $url"
					#$AllSearchResults += Invoke-MSGraphGetRequestWithMSGraphAllPages $url

					$DevicesWhereUserLoggedOn += Invoke-MSGraphGetRequestWithMSGraphAllPages $url
					
					# Sort by lastLogOnDateTime (newest first)
					#$DevicesWhereUserLoggedOn = $DevicesWhereUserLoggedOn | Sort-Object { lastLogOnDateTime -Descending
					
					# Flatten objects and add new properties for ComboBox results
					foreach($device in $DevicesWhereUserLoggedOn) {
						
						$userLoggedOnUserId = $device.usersLoggedOn | Where-Object -Property userId -eq $AzureADUser.id | Select-Object -ExpandProperty userId
						
						# There can be multiple lastLogOnDateTime properties
						# Sort Descending and select first to get latest date
						$userLoggedOnlastLogOnDateTime = $device.usersLoggedOn | Where-Object -Property userId -eq $AzureADUser.id | Sort-Object -Property lastLogOnDateTime -Descending | Select-Object -First 1 -ExpandProperty lastLogOnDateTime
						
						# Add new attributes to object
						$device | Add-Member -MemberType noteProperty -Name userLoggedOnUserId -Value $userLoggedOnUserId
						
						$device | Add-Member -MemberType noteProperty -Name userLoggedOnlastLogOnDateTime -Value $userLoggedOnlastLogOnDateTime
						
						$userLoggedOnlastLogOnDateTimeDaysAgo = New-TimeSpan (Get-Date) (Get-Date $userLoggedOnlastLogOnDateTime) | Select-Object -ExpandProperty Days
						$device | Add-Member -MemberType noteProperty -Name userLoggedOnlastLogOnDateTimeDaysAgo -Value $userLoggedOnlastLogOnDateTimeDaysAgo

						# Add new attribute to object
						
						#$searchStringDeviceProperty = "$($device.deviceName)`t`t(Logged-On $userLoggedOnlastLogOnDateTimeDaysAgo days ago)"
						
						$searchStringDeviceProperty = "{0,-20} {1,-8} {2,4} {3,8}" -f $device.deviceName, "Logged-On", $userLoggedOnlastLogOnDateTimeDaysAgo, "days ago"
						
						$device | Add-Member -MemberType noteProperty -Name searchStringDeviceProperty -Value $searchStringDeviceProperty
						
					}
					
					$AllSearchResults += $DevicesWhereUserLoggedOn | Sort-Object -Property userLoggedOnlastLogOnDateTime -Descending

				}
			}
		}	
	}
	
	# Make sure our results are in array
	$AllSearchResults = [array]$AllSearchResults
	
	# DEBUG
	#Write-Verbose "AllSearchResults: $AllSearchResults"

	if($AllSearchResults) {
		
		# Create SearchResultToolTip information for each device
		Foreach($device in $AllSearchResults) {
			$ToolTip = ($device | Select-Object -Property deviceName,userPrincipalName,operatingSystem,Manufacturer,Model,SerialNumber | Format-List | Out-String).Trim()
			
			$device | Add-Member -MemberType noteProperty -Name SearchResultToolTip -Value $ToolTip
		}

		Write-Verbose "Found $($AllSearchResults.Count) devices"
		#Write-Verbose "$($AllSearchResults.deviceName -join "`n")"

		$script:ComboItemsInSearchComboBoxSource = $AllSearchResults

		# Set items source object array
		# Clear CreateReport ComboBox
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.Items.Clear()
		
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.ItemsSource = [array]$script:ComboItemsInSearchComboBoxSource
		
		# Specify what property to show in combobox dropdown list
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.DisplayMemberPath = 'searchStringDeviceProperty'

		# preselect the first element
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.SelectedIndex = 0
		
		# Enable Create report button
		$WPFButton_GridIntuneDeviceDetailsBorderTop_CreateReport.isEnabled = $True
		
		# Set button font to bold
		$WPFButton_GridIntuneDeviceDetailsBorderTop_CreateReport.FontWeight = "Bold"

		if($AllSearchResults.Count -gt 1) {
			# Open dropdown list if we got more than 1 devices from search
			$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.IsDropDownOpen = $true
			
			Write-host "Found $($AllSearchResults.Count) devices"
			$WPFlabel_GridIntuneDeviceDetailsBorderTop_FoundXDevices.Content = "Found $($AllSearchResults.Count) devices"
			$WPFlabel_GridIntuneDeviceDetailsBorderTop_FoundXDevices.Visibility = 'Visible'
			
		} else {
			# Close dropdown list
			#$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.IsDropDownOpen = $false
			$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.IsDropDownOpen = $true
			
			$WPFlabel_GridIntuneDeviceDetailsBorderTop_FoundXDevices.Content = "Found $($AllSearchResults.Count) device"
			$WPFlabel_GridIntuneDeviceDetailsBorderTop_FoundXDevices.Visibility = 'Visible'
			Write-host "Found $($AllSearchResults.Count) device: $($AllSearchResults.deviceName)"
		}


	} else {
		# Did not find any devices
		Write-Host "Found 0 devices"
		$WPFlabel_GridIntuneDeviceDetailsBorderTop_FoundXDevices.Content = "Did not find devices"
		$WPFlabel_GridIntuneDeviceDetailsBorderTop_FoundXDevices.Visibility = 'Visible'
		$WPFIntuneDeviceDetails_textBox_DeviceName.Text = "Could not find device"
		$WPFIntuneDeviceDetails_textBox_DeviceName.Foreground = "red"
	}
	
}



# Main function to gather data to UI
function Get-DeviceInformation {
	Param(
	    [Parameter(Mandatory=$true,
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
		
		[Parameter(Mandatory=$false)]
		[boolean]$ReloadCacheData=$false
	)

	$UnknownAssignmentGroupFound = $false

	# Clear variables
	$Script:IntuneManagedDevice = $null
	$Script:AutopilotDeviceWithAutpilotProfile = $null
	$Script:EnrollmentStatusPagePolicyId = $null
	$Script:EnrollmentStatusPageProfileName = $null
	$Script:PrimaryUserGroupsMemberOf = $null
	$Script:PrimaryUser = $null
	$Script:AzureADDevice = $null
	$Script:deviceGroupMemberships = $null
	$Script:LatestCheckedinUser = $null
	
	# Prepare Application Assignments user select dropdown combobox source
	# Add device (no primary user), primary user and logged on users to list
	# Pre-Select either Primary User or device depending if there is primary user
	$DeviceWithoutUserObject = [PSCustomObject]@{
		primaryUserId     = '00000000-0000-0000-0000-000000000000'
		id 				  = '00000000-0000-0000-0000-000000000000'
		UserId 			  = '00000000-0000-0000-0000-000000000000'
		userPrincipalName = 'Device without user'
	}

	# Create array for Select user ComboBox source objects
	$script:ComboItemsInApplicationSelectUserComboBoxSource = @()
	
	# Add user object to Application Assignments user select dropdown combobox source
	$script:ComboItemsInApplicationSelectUserComboBoxSource += $DeviceWithoutUserObject
		
	$IntuneDeviceId = $id
	
	# Get Intune device object
	$Script:IntuneManagedDevice = Get-IntuneManagedDevice -managedDeviceId $IntuneDeviceId
	$Success = $?

	if (-not $Success) {
		Write-Host "Error finding Intune deviceId $IntuneDeviceId!" -ForegroundColor Red
		$WPFIntuneDeviceDetails_textBox_DeviceName.Text = "Could not find device"
        $WPFIntuneDeviceDetails_textBox_DeviceName.Foreground = "red"
		return
	}

	Write-Host
	Write-Host "Creating report for device: $($Script:IntuneManagedDevice.deviceName)"

	# DEBUG
	# Copy json data to clipboard so you can paste json to any text editor
	#$Script:IntuneConfigurationProfilesWithAssignments | ConvertTo-Json -Depth 6 | Set-Clipboard

	Write-Host "Get additional user and device data"

    $WPFIntuneDeviceDetails_textBox_DeviceName.Text = $Script:IntuneManagedDevice.DeviceName
	$WPFIntuneDeviceDetails_textBox_DeviceName.Foreground = "#FF004CFF"
	
	# Add ToolTip info to deviceName TextBox
	$WPFIntuneDeviceDetails_textBox_DeviceName_ToolTip_DeviceName.Text = $Script:IntuneManagedDevice.DeviceName
	$DeviceNameToolTip = ($Script:IntuneManagedDevice | Select-Object -Property userPrincipalName,operatingSystem,osVersion,ownerType,deviceType,Manufacturer,Model,chassisType,serialNumber,deviceEnrollmentType,joinType,managedDeviceName,autopilotEnrolled,enrollmentProfileName,enrolledDateTime | Format-List | Out-String).Trim()
	$WPFIntuneDeviceDetails_textBox_DeviceName_ToolTip_DeviceProperties.Text = $DeviceNameToolTip
	
	# Overview tab Device information
	$OverviewTAB_DeviceInformation_Text = [array]$null
	$WPFTextBlock_Overview_DeviceName.Text = "Device Name: $($Script:IntuneManagedDevice.deviceName)"
	
	$OverviewDeviceInformation = ($Script:IntuneManagedDevice | Select-Object -Property userPrincipalName,operatingSystem,osVersion,skuFamily,ownerType,deviceType,Manufacturer,Model,chassisType,serialNumber,joinType,deviceEnrollmentType,managedDeviceName,autopilotEnrolled,enrollmentProfileName,enrolledDateTime | Format-List | Out-String).Trim()
	
	$OverviewTAB_DeviceInformation_Text += $OverviewDeviceInformation
	$WPFTextBox_Overview_Device.Text = $OverviewTAB_DeviceInformation_Text
	
	# Enable DeviceName right click menus
	$WPFIntuneDeviceDetails_textBox_DeviceName_Menu_OpenDeviceInBrowser.isEnabled = $True
	$WPFIntuneDeviceDetails_textBox_DeviceName_Menu_OpenAzureADDeviceInBrowser.isEnabled = $True
	
	# Enable Compliance right click menu
	$WPFtextBox_Compliance_Menu_OpenDeviceComplianceInBrowser.isEnabled = $True
	
    $WPFManufacturer_textBox.Text = $Script:IntuneManagedDevice.Manufacturer
    $WPFModel_textBox.Text = $Script:IntuneManagedDevice.Model
    $WPFSerial_textBox.Text = $Script:IntuneManagedDevice.serialNumber
    $WPFWiFi_textBox.Text = $Script:IntuneManagedDevice.wiFiMacAddress

	# Not sure if this is still problem (14.4.2022)
    # These does not seem to work when getting only 1 device
    # Getting all devices results ok values
    # Workaround is to use hardwareInformation attribute
    #$Script:IntuneManagedDevice.totalStorageSpaceInBytes
    #$Script:IntuneManagedDevice.freeStorageSpaceInBytes

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
		if($operatingSystemEdition -eq 'Windows 11 SE') { $operatingSystemEdition = 'SE' }

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

    if($Script:IntuneManagedDevice.operatingSystem -eq 'Windows') {
        switch -wildcard ($Script:IntuneManagedDevice.osVersion) {
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
			'10.0.19045.*' {   $Version = '10 22H2'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2025-05-13 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2024-05-14 | Select-Object -ExpandProperty Days
                        }
                    }
			'10.0.22000.*' {   $Version = '11 21H2'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2024-10-08 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*') -or ($operatingSystemEdition -like '*SE*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2023-10-10 | Select-Object -ExpandProperty Days
                        }
                    }
			'10.0.22621.*' {   $Version = '11 22H2'
                        if(($operatingSystemEdition -eq 'Ent') -or ($operatingSystemEdition -eq 'Edu')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2025-10-14 | Select-Object -ExpandProperty Days
                        }
                        if(($operatingSystemEdition -eq 'Home') -or ($operatingSystemEdition -like '*Pro*') -or ($operatingSystemEdition -like '*SE*')) {
                            $WindowsSupportEndsInDays = New-TimeSpan $CurrentDate 2024-10-08 | Select-Object -ExpandProperty Days
                        }
                    }
            Default {
                        $Version = $Script:IntuneManagedDevice.operatingSystem
                    }
        }

        if([double]$WindowsSupportEndsInDays -lt 0) {
            # Windows 10 support already ended
            $OSVersionToolTip = "Version $($Script:IntuneManagedDevice.osVersion)`n`nSupport for this Windows version has already ended $WindowsSupportEndsInDays days ago!`n`nUpdate device immediately!"
            
            # Red background for OSVersion textbox
            $WPFOSVersion_textBox.Background = '#FF6347'
            $WPFOSVersion_textBox.Foreground = '#000000'

        } elseif (([double]$WindowsSupportEndsInDays -ge 0) -and ([double]$WindowsSupportEndsInDays -le 90)) {
            # Windows 10 support is ending in 30 days
            $OSVersionToolTip = "Version $($Script:IntuneManagedDevice.osVersion)`n`nSupport for this Windows version is ending in $WindowsSupportEndsInDays days.`n`nSchedule Windows upgrade for this device."
            
            # Yellow background for OSVersion textbox
            $WPFOSVersion_textBox.Background = 'yellow'
            $WPFOSVersion_textBox.Foreground = '#000000'
            
        } elseif([double]$WindowsSupportEndsInDays -gt 90) {
            # Windows 10 has support over 30 days
            $OSVersionToolTip = "Version $($Script:IntuneManagedDevice.osVersion)`n`nSupport for this Windows version will end in $WindowsSupportEndsInDays days."

            # Green background for OSVersion textbox
            $WPFOSVersion_textBox.Background = '#7FFF00'
            $WPFOSVersion_textBox.Foreground = '#000000'
        }
        $WPFOSVersion_textBox.Tooltip = $OSVersionToolTip

    } else {
        $Version = $Script:IntuneManagedDevice.osVersion
    }
    $WPFOSVersion_textBox.Text = "$($Script:IntuneManagedDevice.operatingSystem) $Version $operatingSystemEdition"

    $WPFCompliance_textBox.Text = $Script:IntuneManagedDevice.complianceState
    if($Script:IntuneManagedDevice.complianceState -eq 'compliant') {
        $WPFCompliance_textBox.Background = '#7FFF00'
        $WPFCompliance_textBox.Foreground = '#000000'
    }

    if($Script:IntuneManagedDevice.complianceState -eq 'noncompliant') {
        $WPFCompliance_textBox.Background = '#FF6347'
        $WPFCompliance_textBox.Foreground = '#000000'
    }

    if($Script:IntuneManagedDevice.complianceState -eq 'unknown') {
        $WPFCompliance_textBox.Background = 'yellow'
        $WPFCompliance_textBox.Foreground = '#000000'
    }

    $WPFisEncrypted_textBox.Text = $Script:IntuneManagedDevice.isEncrypted
    if($Script:IntuneManagedDevice.isEncrypted -eq 'True') {
        $WPFisEncrypted_textBox.Background = '#7FFF00'
        $WPFisEncrypted_textBox.Foreground = '#000000'
    } else {
        $WPFisEncrypted_textBox.Background = '#FF6347'
        $WPFisEncrypted_textBox.Foreground = '#000000'
    }


	# Change DateTime to local timezone
	$lastSyncDateTimeLocalTime = (Get-Date $Script:IntuneManagedDevice.lastSyncDateTime).ToLocalTime()
    
    # This is used in textBox ToolTip
    $lastSyncDateTimeUFormatted = Get-Date $lastSyncDateTimeLocalTime -uformat '%G-%m-%d %H:%M:%S'
	
    $WPFlastSync_textBox.Tooltip = "Last Sync DateTime (yyyy-MM-dd HH:mm:ss): $lastSyncDateTimeUFormatted"

    $lastSyncDays = (New-Timespan $lastSyncDateTimeLocalTime).Days
    $lastSyncHours = (New-Timespan $lastSyncDateTimeLocalTime).Hours
   
    # This would be in UTC time with this syntax
    #$enrolledDays = (New-Timespan $Script:IntuneManagedDevice.enrolledDateTime).Days

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
    $WPFAutopilotEnrolled_textBox.Text = $Script:IntuneManagedDevice.autopilotEnrolled

	# Disable MenuItem Open Autopilot Devices...
	# We will enable these below if Device autopilot info is found
	$WPFIntuneDeviceDetails_textBox_DeviceName_Menu_OpenAutopilotDeviceInBrowser.isEnabled = $False
	$WPFAutopilotEnrolled_textBox_Menu_OpenAutopilotDeviceInBrowser.isEnabled = $False
	$WPFAutopilotProfile_textBox_Menu_OpenAutopilotDeploymentProfileInBrowser.isEnabled = $False

    # Get Windows Autopilot GroupTag and assigned profile if device is autopilotEnrolled
    if($Script:IntuneManagedDevice.autopilotEnrolled) {
        # Find device from Windows Autopilot
        $url = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,%27$($Script:IntuneManagedDevice.serialNumber)%27)&_=1577625591868"
        $AutopilotDevice = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

        if($AutopilotDevice) {

			# Set Autopilot Device JSON information to right side tabItem textBox
			$WPFAutopilotDevice_json_textBox.Text = $AutopilotDevice | ConvertTo-Json -Depth 5

			# Enable MenuItem Open Autopilot Devices...
			$WPFIntuneDeviceDetails_textBox_DeviceName_Menu_OpenAutopilotDeviceInBrowser.isEnabled = $True
			$WPFAutopilotEnrolled_textBox_Menu_OpenAutopilotDeviceInBrowser.isEnabled = $True

            # Get more detailed information including Windows Autopilot IntendedProfile
            $url = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities/$($AutopilotDevice.id)?`$expand=deploymentProfile,intendedDeploymentProfile&_=1578315612557"
            $Script:AutopilotDeviceWithAutpilotProfile = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

            if($Script:AutopilotDeviceWithAutpilotProfile) {
				# Enable MenuItem Open Autopilot Deployment Profile in browser
				$WPFAutopilotProfile_textBox_Menu_OpenAutopilotDeploymentProfileInBrowser.isEnabled = $True

                $WPFAutopilotGroupTag_textBox.Text = $Script:AutopilotDeviceWithAutpilotProfile.groupTag
                
                # Check that assigned Autopilot Profile is same than Intended Autopilot Profile
                if($Script:AutopilotDeviceWithAutpilotProfile.deploymentProfile.id -eq $Script:AutopilotDeviceWithAutpilotProfile.intendedDeploymentProfile.id) {
                    $WPFAutopilotProfile_textBox.Text = $Script:AutopilotDeviceWithAutpilotProfile.deploymentProfile.displayName
                } else {
                    # Windows Autopilot current and intended profile are different so we have to wait Autopilot sync to fix that
                    $WPFAutopilotProfile_textBox.Text = 'Profile sync is not ready'
                }

				# Removed these tooltips because there is separate JSON view for AutopilotProfile JSON
				#$WPFAutopilotProfile_textBox.Tooltip = $Script:AutopilotDeviceWithAutpilotProfile | ConvertTo-Json
                #$WPFAutopilotGroupTag_textBox.Tooltip = $Script:AutopilotDeviceWithAutpilotProfile | ConvertTo-Json
				
				# Set Autopilot Enrollment Profile to JSON view
				$WPFAutopilotEnrollmentProfile_json_textBox.Text = $Script:AutopilotDeviceWithAutpilotProfile | ConvertTo-Json -Depth 4
            }
        }
    }

	# Get EnrollmentConfigurationPolicies
	# Get Enrollment Status Page and Enrollment Restrictions info
	# Note this is POST type request
	$url = 'https://graph.microsoft.com/beta/deviceManagement/reports/getEnrollmentConfigurationPoliciesByDevice'

	$GraphAPIPostRequest = @"
{
    "select": [
        "PolicyId",
        "PolicyType",
        "ProfileName"
    ],
    "filter": "(DeviceId eq '$IntuneDeviceId')"
}
"@

	Write-Host "Get Intune Enrollment Status Page and Enrollment Restrictions used in this device's enrollment"
	$MSGraphRequest = Invoke-MSGraphRequest -Url $url -Content $GraphAPIPostRequest.ToString() -HttpMethod 'POST'
	$Success = $?

	if($Success) {
		#Write-Host "Success"
	} else {
		# Invoke-MSGraphRequest failed
		Write-Error "Error getting Intune Enrollment Status Page and Enrollment Restrictions information"
		return 1
	}

	# Get AllMSGraph pages
	# This is also workaround to get objects without assigning them from .Value attribute
	$EnrollmentConfigurationPoliciesByDevice = Get-MSGraphAllPages -SearchResult $MSGraphRequest
	$Success = $?

	if($Success) {
		#Write-Host "Success"
	} else {
		# Invoke-MSGraphRequest failed
		Write-Error "Error getting Intune Enrollment Status Page and Enrollment Restrictions information"
		return 1
	}

	$EnrollmentConfigurationPoliciesByDevice = Objectify_JSON_Schema_and_Data_To_PowershellObjects $EnrollmentConfigurationPoliciesByDevice

	# DEBUG
	#$EnrollmentConfigurationPoliciesByDevice | ConvertTo-Json | Set-Clipboard
	
	$DeviceTypeEnrollmentRestrictionObject = $EnrollmentConfigurationPoliciesByDevice | Where-Object PolicyType_loc -eq 'Device type enrollment restriction'
	$DeviceTypeEnrollmentRestrictionPolicyId = $DeviceTypeEnrollmentRestrictionObject.PolicyId
	$DeviceTypeEnrollmentRestrictionProfileName = $DeviceTypeEnrollmentRestrictionObject.ProfileName
	$WPFtextBox_EnrollmentRestrictions.Text = $DeviceTypeEnrollmentRestrictionProfileName

	if($DeviceTypeEnrollmentRestrictionPolicyId) {
		# Enable right click menu if there is Device Type Enrollment Restriction Policy Id
		$WPFtextBox_EnrollmentRestrictions_Menu_OpenEnrollmentRestrictionProfilesInBrowser.isEnabled = $True
	}



	$EnrollmentStatusPageObject = $EnrollmentConfigurationPoliciesByDevice | Where-Object PolicyType_loc -eq 'Enrollment status page'
	$Script:EnrollmentStatusPagePolicyId = $EnrollmentStatusPageObject.PolicyId
	$Script:EnrollmentStatusPageProfileName = $EnrollmentStatusPageObject.ProfileName
	$WPFtextBox_EnrollmentStatusPage.Text = $EnrollmentStatusPageProfileName

	if($Script:EnrollmentStatusPagePolicyId) {
		# Enable right click menu if there is EnrollmentProfile ID
		$WPFtextBox_EnrollmentStatusPage_Menu_OpenESPProfileInBrowser.isEnabled = $True
	}

	
	########
    # Get Intune device primary user
    $url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($IntuneDeviceId)/users"
    $primaryUserId = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
    $primaryUserId = $primaryUserId.id

    # Set Primary User id to zero if we got nothing. This is used later with Application assignment report information
    if($primaryUserId -eq $null) {
        $primaryUserId = '00000000-0000-0000-0000-000000000000'
		
		# Set PrimaryUser MenuItem DISABLED because we did NOT find user
		$WPFprimaryUser_textBox_Menu_OpenPrimaryUserInBrowser.isEnabled = $False
    }

    $Script:PrimaryUserGroupsMemberOf = $null

	$PrimaryUserMail = $null
	$PrimaryUserUPN = $null
    if($primaryUserId -ne '00000000-0000-0000-0000-000000000000') {
    
        $Script:PrimaryUser = $null

        # Check we have valid GUID
        if([System.Guid]::Parse($primaryUserId)) {
			# Get all user data so we have full information in JSON view also
			$url = "https://graph.microsoft.com/beta/users/$($primaryUserId)"
            $Script:PrimaryUser = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
        }

        if($Script:PrimaryUser) {

			# Add user object to Application Assignments user select dropdown combobox source
			# Create a COPY of object so we can change it's values without making change to original object
			$script:ComboItemsInApplicationSelectUserComboBoxSource += $Script:PrimaryUser.PsObject.Copy()
			
			# Edit ComboItemsInApplicationSelectUserComboBoxSource PrimaryUser UserPrincipalName to (Primary User) UserPrincipalName
			# $variable[-1] refers to last item in array which is latest added Primary User to our array
			$script:ComboItemsInApplicationSelectUserComboBoxSource[-1].UserPrincipalName = "$($Script:PrimaryUser.UserPrincipalName) (Primary User)"

			# Set PrimaryUser MenuItem ENABLED because we found user
			$WPFprimaryUser_textBox_Menu_OpenPrimaryUserInBrowser.isEnabled = $True

			$PrimaryUserToolTip = $null
			
			$PrimaryUserMail = $Script:PrimaryUser.mail
			$PrimaryUserUserPrincipalName = $Script:PrimaryUser.userPrincipalName
            $WPFprimaryUser_textBox.Text = $PrimaryUserUserPrincipalName

			# Add info to PrimaryUser Tooltip
			$WPFtextBlock_primaryUser_textBox_ToolTip_UPN.Text = $PrimaryUserUserPrincipalName
			
			$BasicInfo = ($Script:PrimaryUser | Select-Object -Property accountEnabled,displayName,userPrincipalName,email,userType,mobilePhone,jobTitle,department,companyName,employeeId,employeeType,streetAddress,postalCode,state,country,officeLocation,usageLocation | Format-List | Out-String).Trim()
			$WPFtextBlock_primaryUser_textBox_ToolTip_BasicInfo.Text = $BasicInfo
			
			$proxyAddressesToolTip = ($Script:PrimaryUser.proxyAddresses | Format-List | Out-String).Trim()
			$WPFtextBlock_primaryUser_textBox_ToolTip_proxyAddresses.Text = $proxyAddressesToolTip

			$otherMailsToolTip = ($Script:PrimaryUser.otherMails | Format-List | Out-String).Trim()
			$WPFtextBlock_primaryUser_textBox_ToolTip_otherMails.Text = $otherMailsToolTip

			$onPremisesAttributesToolTip = ($Script:PrimaryUser | Select-Object -Property onPremisesSamAccountName, onPremisesUserPrincipalName, onPremisesSyncEnabled, onPremisesLastSyncDateTime, onPremisesDomainName, onPremisesDistinguishedName,onPremisesImmutableId | Format-List | Out-String).Trim()
			$WPFtextBlock_primaryUser_textBox_ToolTip_onPremisesAttributes.Text = $onPremisesAttributesToolTip

			$onPremisesExtensionAttributesToolTip = ($Script:PrimaryUser.onPremisesExtensionAttributes | Format-List | Out-String).Trim()
			$WPFtextBlock_primaryUser_textBox_ToolTip_onPremisesExtensionAttributes.Text = $onPremisesExtensionAttributesToolTip
			
			# Overview tab Device information
			$OverviewTAB_PrimaryUserInformation_Text = [array]$null
			$WPFTextBlock_Overview_PrimaryUserName.Text = "Primary User: $PrimaryUserUserPrincipalName"
			$OverviewTAB_PrimaryUserInformation_Text += $BasicInfo
			$WPFTextBox_Overview_PrimaryUser.Text = $OverviewTAB_PrimaryUserInformation_Text
			
			
			
			# Set PrimaryUser information to JSON view
			$WPFPrimaryUserDetails_json_textBox.Text = $Script:PrimaryUser | ConvertTo-Json -Depth 4

			# Get Primary User Groups memberOf
            $url = "https://graph.microsoft.com/beta/users/$($primaryUserId)/memberOf?_=1577625591876"
            $Script:PrimaryUserGroupsMemberOf = Invoke-MSGraphGetRequestWithMSGraphAllPages $url
            if($Script:PrimaryUserGroupsMemberOf) {

				# DEBUG $UserMemberOfGroupNames -> Paste json data to text editor
				#$Script:PrimaryUserGroupsMemberOf | ConvertTo-Json -Depth 4 | Set-ClipBoard

				$Script:PrimaryUserGroupsMemberOf = Add-AzureADGroupGroupTypeExtraProperties $Script:PrimaryUserGroupsMemberOf
				
				Write-Verbose "Add Primary User's AzureAD groups devices and users member count custom properties"
				$Script:PrimaryUserGroupsMemberOf = Add-AzureADGroupDevicesAndUserMemberCountExtraProperties $Script:PrimaryUserGroupsMemberOf


				$script:PrimaryUserGroupMembershipsObservableCollection = $Script:PrimaryUserGroupsMemberOf | Sort-Object -Property displayName

				if($script:PrimaryUserGroupMembershipsObservableCollection.Count -gt 1) {
					# ItemsSource works if we are sorting 2 or more objects
					$WPFlistView_PrimaryUser_GroupMemberships.Itemssource = $script:PrimaryUserGroupMembershipsObservableCollection | Sort-Object -Property displayName
				} else {
					# Only 1 object so we can't do sorting
					# If we try to sort here then our object array breaks and it does not work for ItemsSource
					# Cast as array because otherwise it will fail
					$WPFlistView_PrimaryUser_GroupMemberships.Itemssource = [array]$script:PrimaryUserGroupMembershipsObservableCollection
				}
   
				# Enable Primary User Group Memberships right click menus
				$WPFListView_GridTabItem_PrimaryUser_GroupMembershipsTAB_Menu_Copy_DynamicRules.isEnabled = $True
				$WPFListView_GridTabItem_PrimaryUser_GroupMembershipsTAB_Menu_Copy_JSON.isEnabled = $True
				$WPFListView_GridTabItem_PrimaryUser_GroupMembershipsTAB_Menu_Open_Group_In_Browser.isEnabled = $True
  

				# Set ToolTip to TabItem Header showing all Azure AD Groups
				$PrimaryUserGroupsMemberOfToolTip = [array]$null
				$Script:PrimaryUserGroupsMemberOf | Sort-Object -Property displayName | Foreach { $PrimaryUserGroupsMemberOfToolTip += "$($_.displayName)`n" }
				$WPFTextBlock_TabItem_User_GroupMembershipsTAB_Header_ToolTip.Text = $PrimaryUserGroupsMemberOfToolTip
  
            } else {
                Write-Host "Did not find any groups for user $($primaryUserId)"
            }
        } else {
            $WPFprimaryUser_textBox.Text = 'user not found'
        }
    } else {
        $WPFprimaryUser_textBox.Text = 'Shared device'
		$WPFTextBlock_Overview_PrimaryUserName.Text = 'No Primary User'

		# Make sure PrimaryUser ToolTip information is cleared
		$WPFtextBlock_primaryUser_textBox_ToolTip_UPN.Text = ''
		$WPFtextBlock_primaryUser_textBox_ToolTip_BasicInfo.Text = ''
		$WPFtextBlock_primaryUser_textBox_ToolTip_proxyAddresses.Text = ''
		$WPFtextBlock_primaryUser_textBox_ToolTip_otherMails.Text = ''
		$WPFtextBlock_primaryUser_textBox_ToolTip_onPremisesAttributes.Text = ''
		$WPFtextBlock_primaryUser_textBox_ToolTip_onPremisesExtensionAttributes.Text = ''
    }


	# Get Logged on users information
	Get-CheckedInUsersGroupMemberships

    # Get Device AzureADGroup memberships
    $url = "https://graph.microsoft.com/beta/devices?`$filter=deviceid%20eq%20`'$($Script:IntuneManagedDevice.azureADDeviceId)`'"
    $Script:AzureADDevice = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

    if($Script:AzureADDevice) {

		# Set info to Azure AD Device JSON view
		$WPFAzureADDeviceDetails_json_textBox.Text = $Script:AzureADDevice | ConvertTo-Json -Depth 4

		# DeviceName ToolTip extensionAttributes
		$AzureADDeviceExtensionAttributesToolTip = ($Script:AzureADDevice.extensionAttributes | Format-List | Out-String).Trim()
		$WPFIntuneDeviceDetails_textBox_DeviceName_ToolTip_extensionAttributes.Text = $AzureADDeviceExtensionAttributesToolTip

        $url = "https://graph.microsoft.com/beta/devices/$($Script:AzureADDevice.id)/memberOf?_=1577625591876"
        $Script:deviceGroupMemberships = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

        if($Script:deviceGroupMemberships) {

			# DEBUG $Script:deviceGroupMemberships -> Paste json data to text editor
			#$Script:deviceGroupMemberships | ConvertTo-Json -Depth 4 | Set-ClipBoard

			# Add extra properties like GroupMembers, GroupType, Security and MembershipType
			$Script:deviceGroupMemberships = Add-AzureADGroupGroupTypeExtraProperties $Script:deviceGroupMemberships
			
			Write-Verbose "Add device's AzureAD groups devices and users member count custom properties"
			$Script:deviceGroupMemberships = Add-AzureADGroupDevicesAndUserMemberCountExtraProperties $Script:deviceGroupMemberships

			$script:DeviceGroupMembershipsObservableCollection = $Script:deviceGroupMemberships | Sort-Object -Property displayName

			if($script:DeviceGroupMembershipsObservableCollection.Count -gt 1) {
				# ItemsSource works if we are sorting 2 or more objects
				$WPFlistView_Device_GroupMemberships.Itemssource = $script:DeviceGroupMembershipsObservableCollection | Sort-Object -Property displayName
			} else {
				# Only 1 object so we can't do sorting
				# If we try to sort here then our object array breaks and it does not work for ItemsSource
				# Cast as array because otherwise it will fail
				$WPFlistView_Device_GroupMemberships.Itemssource = [array]$script:DeviceGroupMembershipsObservableCollection
			}
			
			# Enable Device Group Memberships right click menus
			$WPFListView_GridTabItem_Device_GroupMembershipsTAB_Menu_Copy_DynamicRules.isEnabled = $True
			$WPFListView_GridTabItem_Device_GroupMembershipsTAB_Menu_Copy_JSON.isEnabled = $True
			$WPFListView_GridTabItem_Device_GroupMembershipsTAB_Menu_Open_Group_In_Browser.isEnabled = $True
			
			# Set ToolTip to TabItem Header showing all Azure AD Groups
			$DeviceGroupsMemberOfToolTip = [array]$null
			$Script:deviceGroupMemberships | Sort-Object -Property displayName | Foreach { $DeviceGroupsMemberOfToolTip += "$($_.displayName)`n" }
			
			#$WPFTabItem_Device_GroupMemberships_Header.ToolTip = $DeviceGroupsMemberOfToolTip
			$WPFTextBlock_TabItem_Device_GroupMemberships_Header_ToolTip.Text = $DeviceGroupsMemberOfToolTip
			
			
			
        } else {
            #Write-Host "Did not find any device groupmemberships for"
        }
    }

    # Device json information

    # Copy additional information attributes
    $Script:IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name activationLockBypassCode -Value "$($AdditionalDeviceInformation.activationLockBypassCode)" -Force
    $Script:IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name iccid -Value "$($AdditionalDeviceInformation.iccid)" -Force
    $Script:IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name udid -Value "$($AdditionalDeviceInformation.udid)" -Force
    $Script:IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name roleScopeTagIds -Value "$($AdditionalDeviceInformation.roleScopeTagIds)" -Force
    $Script:IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name ethernetMacAddress -Value "$($AdditionalDeviceInformation.ethernetMacAddress)" -Force
    $Script:IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name processorArchitecture -Value "$($AdditionalDeviceInformation.processorArchitecture)" -Force
    $Script:IntuneManagedDevice | Add-Member -MemberType NoteProperty -Name hardwareInformation -Value "$($AdditionalDeviceInformation.hardwareInformation)" -Force

    $WPFIntuneDeviceDetails_json_textBox.Text = $Script:IntuneManagedDevice | ConvertTo-Json -Depth 5
	
	Write-Host "Success"
	Write-Host


	# Skip Applications and Configurations report if checkBox is checked
	# CheckBox: Basic info only (Skip Apps and Configurations)
	if(-not $WPFcheckBox_SkipAppAndConfigurationAssignmentReport.IsChecked) {

		Write-Host
		Write-Host "Download Intune filters"
		$AllIntuneFilters = Download-IntuneFilters
		Write-Host "Found $($AllIntuneFilters.Count) filters"
		Write-Host

	<#
		# NOTE Intune Filters do not support Graph API filtering with lastModifiedDateTime
		# So we have to get filters every time
		# This is left here for future purposes if we get lastModifiedDateTime support in the future
		
		Write-Host
		Write-Host "Download filters"
		#$AllIntuneFilters = Download-IntuneFilters
		$AllIntuneFilters = Download-IntuneConfigurationProfiles2 'https://graph.microsoft.com/beta/deviceManagement/assignmentFilters?$select=*' 'filters.json'
		Write-Host "Found $($AllIntuneFilters.Count) filters"
		Write-Host
	#>



		# Get Intune Application Assignments
		Write-Host "Get Intune Applications"
		$Script:AppsWithAssignments = Get-ApplicationsWithAssignments -ReloadCacheData $ReloadCacheData
		Write-Host "Found $($Script:AppsWithAssignments.Count) applications"
		Write-Host


		# Set items source object array to Selected User combobox
		$WPFIntuneDeviceDetails_ApplicationAssignments_SelectUser_ComboBox.ItemsSource = [array]$script:ComboItemsInApplicationSelectUserComboBoxSource
		
		# Specify what property to show in Selected User combobox dropdown list
		$WPFIntuneDeviceDetails_ApplicationAssignments_SelectUser_ComboBox.DisplayMemberPath = 'UserPrincipalName'


		if($Script:PrimaryUser) {
			# preselect the second element which is hardcoded as Primary User
			$WPFIntuneDeviceDetails_ApplicationAssignments_SelectUser_ComboBox.SelectedIndex = 1
		} else {
			# preselect the first element which is hardcoded as Device without user
			$WPFIntuneDeviceDetails_ApplicationAssignments_SelectUser_ComboBox.SelectedIndex = 0
		}

		Write-Host "Get Intune Apps and Intents for user"
		$returnMobileAppIntentsForSpecifiedUser = Get-MobileAppIntentsForSpecifiedUser -UserId $primaryUserId -IntuneDeviceId $IntuneDeviceId
		if($returnMobileAppIntentsForSpecifiedUser) {
			Write-Host "Found $returnMobileAppIntentsForSpecifiedUser App Intent and States"
		} else {
			Write-Host "Did not find any App Intent and States"
		}
		Write-Host

		###########################################################################################
		# Create Configuration Assignment information

		Write-Host
		Write-Host "Download Intune configuration profiles"
		
		# Specify variable as array. Otherwise += will not work
		$Script:IntuneConfigurationProfilesWithAssignments = @()

		# User Powershell splatting to specify function parameters
			# Limited properties
			#GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?expand=assignments&$select=id,description,createdDateTime,lastModifiedDateTime,name,assignments'
		$Params = @{
			GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?expand=assignments&$select=*'
			jsonCacheFileName = 'configurationPolicies.json'
			ReloadCacheData = $ReloadCacheData
		}
		$Script:IntuneConfigurationProfilesWithAssignments += Download-IntuneConfigurationProfiles2 @Params

			# Limited properties
			#GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?expand=assignments&$select=id,description,createdDateTime,lastModifiedDateTime,displayname,assignments'
		$Params = @{
			GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?expand=assignments&$select=*'
			jsonCacheFileName = 'groupPolicyConfigurations.json'
			ReloadCacheData = $ReloadCacheData
		}
		$Script:IntuneConfigurationProfilesWithAssignments += Download-IntuneConfigurationProfiles2 @Params

			# Limited properties
			#$GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?expand=assignments&$select=id,description,createdDateTime,lastModifiedDateTime,displayname,assignments'
		$Params = @{
			GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?expand=assignments&$select=*'
			jsonCacheFileName = 'deviceConfigurations.json'
			ReloadCacheData = $ReloadCacheData
		}
		$Script:IntuneConfigurationProfilesWithAssignments += Download-IntuneConfigurationProfiles2 @Params

			# Limited properties
			#raphAPIUrl = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations?expand=assignments&$select=id,description,createdDateTime,lastModifiedDateTime,displayname,assignments'
		$Params = @{
			GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations?expand=assignments&$select=*'
			jsonCacheFileName = 'mobileAppConfigurations.json'
			ReloadCacheData = $ReloadCacheData
		}
		$Script:IntuneConfigurationProfilesWithAssignments += Download-IntuneConfigurationProfiles2 @Params


		$Params = @{
			GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/intents?select=*'
			jsonCacheFileName = 'intents.json'
			ReloadCacheData = $ReloadCacheData
		}
		$Script:IntuneConfigurationProfilesWithAssignments += Download-IntuneConfigurationProfiles2 @Params


		Write-Host "Found $($Script:IntuneConfigurationProfilesWithAssignments.Count) configuration profiles"
		Write-Host

		$script:ConfigurationsAssignmentsObservableCollection = @()

		# Get Device Configurations install state - OLD url
		#$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$IntuneDeviceId/deviceConfigurationStates"

		# Intune uses this Graph API url - OLD url
		#$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$IntuneDeviceId/deviceConfigurationStates?`$filter=((platformType%20eq%20'android')%20or%20(platformType%20eq%20'androidforwork')%20or%20(platformType%20eq%20'androidworkprofile')%20or%20(platformType%20eq%20'ios')%20or%20(platformType%20eq%20'macos')%20or%20(platformType%20eq%20'WindowsPhone81')%20or%20(platformType%20eq%20'Windows81AndLater')%20or%20(platformType%20eq%20'Windows10AndLater')%20or%20(platformType%20eq%20'all'))&_=1578056411936"
		
		# Intune uses this nowdays 30.3.2022
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
		$ConfigurationPoliciesReportForDevice = Download-IntunePostTypeReport -GraphAPIUrl $url -GraphAPIPostRequest $GraphAPIPostRequest
		Write-Host "Found $($ConfigurationPoliciesReportForDevice.Count) Configuration Assignments"


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
			$assignmentGroupId = $null
			$YodamiittiCustomGroupMembers = 'N/A'
			$context = $null
			$DeviceConfiguration = $null
			$IntuneDeviceConfigurationPolicyAssignments = $null
			$IncludeConfigurationAssignmentInSummary = $true
			$properties = $null
			$odatatype = $ConfigurationPolicyReportState.UnifiedPolicyType_loc
			$AssignmentGroupToolTip = $null
			$displayNameToolTip = $null

			$assignmentFilterId = $null
			$assignmentFilterDisplayName = $null
			$FilterToolTip = $null
			$FilterMode = $null


			# Cast as string so our column sorting works
			$YodamiittiCustomMembershipType = [String]''
			
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
				$IntunePolicyObject = $Script:IntuneConfigurationProfilesWithAssignments | Where-Object id -eq $ConfigurationPolicyReportState.PolicyId
				
				$IntuneDeviceConfigurationPolicyAssignments = $IntunePolicyObject.assignments
				$displayNameToolTip = $IntunePolicyObject.description
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
			$odatatype = $odatatype.Replace('#microsoft.graph.', '')
			$assignmentGroup = $null

			foreach ($IntuneDeviceConfigurationPolicyAssignment in $IntuneDeviceConfigurationPolicyAssignments) {

				$assignmentGroup = $null
				$YodamiittiCustomGroupMembers = 'N/A'

				# Only include Configuration which have assignments targeted to this device/user
				$IncludeConfigurationAssignmentInSummary = $false

				$context = '_unknown'
				
				if ($IntuneDeviceConfigurationPolicyAssignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
					# Special case for All Users
					$assignmentGroup = 'All Users'
					$context = 'User'
					$AssignmentGroupToolTip = 'Built-in All Users group'

					$YodamiittiCustomGroupMembers = ''

					$IncludeConfigurationAssignmentInSummary = $true
				}

				if ($IntuneDeviceConfigurationPolicyAssignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
					# Special case for All Devices
					$assignmentGroup = 'All Devices'
					$context = 'Device'
					$AssignmentGroupToolTip = 'Built-in All Devices group'

					$YodamiittiCustomGroupMembers = ''

					$IncludeConfigurationAssignmentInSummary = $true
				}

				if(($IntuneDeviceConfigurationPolicyAssignment.target.'@odata.type' -ne '#microsoft.graph.allLicensedUsersAssignmentTarget') -and ($IntuneDeviceConfigurationPolicyAssignment.target.'@odata.type' -ne '#microsoft.graph.allDevicesAssignmentTarget')) {

					# Group based assignment. We need to get AzureAD Group Name
					# #microsoft.graph.groupAssignmentTarget

					# Test if device is member of this group
					if($Script:deviceGroupMemberships | Where-Object { $_.id -eq $IntuneDeviceConfigurationPolicyAssignment.target.groupId}) {
						
						$assignmentGroupObject = $Script:deviceGroupMemberships | Where-Object { $_.id -eq $IntuneDeviceConfigurationPolicyAssignment.target.groupId}
						
						$assignmentGroup = $assignmentGroupObject.displayName
						$assignmentGroupId = $assignmentGroupObject.id

						# Create Group Members column information
						$DevicesCount = $assignmentGroupObject.YodamiittiCustomGroupMembersCountDevices
						$UsersCount = $assignmentGroupObject.YodamiittiCustomGroupMembersCountUsers
						#$YodamiittiCustomGroupMembers = "$DevicesCount devices, $UsersCount users"
						$YodamiittiCustomGroupMembers = ''
						if($DevicesCount -gt 0) { $YodamiittiCustomGroupMembers += "$DevicesCount devices " }
						if($UsersCount -gt 0) { $YodamiittiCustomGroupMembers += "$UsersCount users " }							

						$AssignmentGroupToolTip = "$($assignmentGroupObject.membershipRule)"
						
						$YodamiittiCustomMembershipType = $assignmentGroupObject.YodamiittiCustomMembershipType
						
						#Write-Host "device group found: $($assignmentGroup.displayName)"
						$context = 'Device'

						$IncludeConfigurationAssignmentInSummary = $true
					} else {
						# Group not found on member of devicegroups
					}

					# Test if primary user is member of assignment group
					if($Script:PrimaryUserGroupsMemberOf | Where-Object { $_.id -eq $IntuneDeviceConfigurationPolicyAssignment.target.groupId}) {
						if($assignmentGroup) {
							# Device also is member of this group. Now we got mixed User and Device memberships
							# Maybe not good practise but it is possible

							# We will actually skip getting possible user Group for this assignment
							# Future improvement is to add user Group information also

							$context = '_Device/User'
						} else {
							# No assignment group was found earlier
							$context = 'User'
						
							$assignmentGroupObject = $Script:PrimaryUserGroupsMemberOf | Where-Object { $_.id -eq $IntuneDeviceConfigurationPolicyAssignment.target.groupId}
							
							$assignmentGroup = $assignmentGroupObject.displayName
							$assignmentGroupId = $assignmentGroupObject.id
							
							# Create Group Members column information
							$DevicesCount = $assignmentGroupObject.YodamiittiCustomGroupMembersCountDevices
							$UsersCount = $assignmentGroupObject.YodamiittiCustomGroupMembersCountUsers
							#$YodamiittiCustomGroupMembers = "$DevicesCount devices, $UsersCount users"
							$YodamiittiCustomGroupMembers = ''
							if($DevicesCount -gt 0) { $YodamiittiCustomGroupMembers += "$DevicesCount devices " }
							if($UsersCount -gt 0) { $YodamiittiCustomGroupMembers += "$UsersCount users " }
							
							$AssignmentGroupToolTip = "$($assignmentGroupObject.membershipRule)"
							
							$YodamiittiCustomMembershipType = $assignmentGroupObject.YodamiittiCustomMembershipType
							
							#Write-Host "User group found: $($assignmentGroup.displayName)"
						}							
						$IncludeConfigurationAssignmentInSummary = $true
					} else {
						# Group not found on member of devicegroups
					}
					
					# Test if Latest LoggedIn User is member of assignment group
					# Only test this if PrimaryUser and Latest LoggedIn User is different user
					if($Script:PrimaryUser.id -ne $Script:LatestCheckedinUser.id) {
						if($Script:LatestCheckedInUserGroupsMemberOf | Where-Object { $_.id -eq $IntuneDeviceConfigurationPolicyAssignment.target.groupId}) {
							if($assignmentGroup) {
								# Device or PrimaryUser also is member of this group.
								# Now we may got mixed User and Device memberships
								# Maybe not good practise but it is possible

								if($context -eq 'Device') {
									$context = '_Device/User'
								}
							} else {
								
								$context = 'User'

								$assignmentGroupObject = $Script:LatestCheckedInUserGroupsMemberOf | Where-Object { $_.id -eq $IntuneDeviceConfigurationPolicyAssignment.target.groupId}
								
								$assignmentGroup = $assignmentGroupObject.displayName
								$assignmentGroupId = $assignmentGroupObject.id
								
								# Create Group Members column information
								$DevicesCount = $assignmentGroupObject.YodamiittiCustomGroupMembersCountDevices
								$UsersCount = $assignmentGroupObject.YodamiittiCustomGroupMembersCountUsers
								$YodamiittiCustomGroupMembers = "$DevicesCount devices, $UsersCount users"
								
								$AssignmentGroupToolTip = "$($assignmentGroupObject.membershipRule)"
								
								$YodamiittiCustomMembershipType = $assignmentGroupObject.YodamiittiCustomMembershipType
								
								#Write-Host "User group found: $($assignmentGroup.displayName)"

								$IncludeConfigurationAssignmentInSummary = $true
							}
						} else {
							# Group not found on member of devicegroups
						}
					}
				}

				
				if($IncludeConfigurationAssignmentInSummary) {
				
					# Set included/excluded attribute
					$PolicyIncludeExclude = ''
					if ($IntuneDeviceConfigurationPolicyAssignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget') {
						$PolicyIncludeExclude = 'Included'
					}
					if ($IntuneDeviceConfigurationPolicyAssignment.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget') {
						$PolicyIncludeExclude = 'Excluded'
					}

					$state = $ConfigurationPolicyReportState.PolicyStatus

					$assignmentFilterId = $IntuneDeviceConfigurationPolicyAssignment.target.deviceAndAppManagementAssignmentFilterId

					#$assignmentFilterDisplayName = $AllIntuneFilters | Where-Object { $_.id -eq $assignmentFilterId } | Select-Object -ExpandProperty displayName
					
					$assignmentFilterObject = $AllIntuneFilters | Where-Object { $_.id -eq $assignmentFilterId }

					$assignmentFilterDisplayName = $assignmentFilterObject.displayName
					$FilterToolTip = $assignmentFilterObject.rule
					
					$FilterMode = $IntuneDeviceConfigurationPolicyAssignment.target.deviceAndAppManagementAssignmentFilterType
					if($FilterMode -eq 'None') {
						$FilterMode = $null
					}

					# Cast variable types to make sure column click based sorting works
					# Sorting may break if there are different kind of objects
					$properties = @{
						context                          = [String]$context
						odatatype                        = [String]$odatatype
						userPrincipalName                = [String]$ConfigurationPolicyReportState.UPN
						displayname                      = [String]$ConfigurationPolicyReportState.PolicyName
						assignmentIntent                 = [String]$assignmentIntent
						IncludeExclude                   = [String]$PolicyIncludeExclude
						assignmentGroup                  = [String]$assignmentGroup
						YodamiittiCustomGroupMembers     = [String]$YodamiittiCustomGroupMembers
						assignmentGroupId 				 = [String]$assignmentGroupId
						state                            = [String]$state
						YodamiittiCustomMembershipType   = [String]$YodamiittiCustomMembershipType
						id                               = $ConfigurationPolicyReportState.PolicyId
						filter							 = [String]$assignmentFilterDisplayName
						filterId						 = [String]$assignmentFilterId
						filterMode						 = [String]$FilterMode
						filterTooltip                    = [String]$FilterTooltip
						AssignmentGroupToolTip 			 = [String]$AssignmentGroupToolTip
						displayNameToolTip               = [String]$displayNameToolTip
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
				# Either assignments does not exists at all
				# or assignment is based on nested groups so earlier check did not find Azure AD group where device and/or user is member

				$context = '_unknown'
				$PolicyIncludeExclude = ''

				# Set variable which we return from this function
				$UnknownAssignmentGroupFound = $true

				# Check if assignments is $null but Policy was found
				# Intune may show Configuration profile status for configuration which is not deployed anymore
				# Check that we did find policy but assignments for that found policy is $null
				if((-not $IntuneDeviceConfigurationPolicyAssignments) -and ($Script:IntuneConfigurationProfilesWithAssignments | Where-Object id -eq $ConfigurationPolicyReportState.PolicyId)) {
					Write-Host "Warning: Policy $($ConfigurationPolicyReportState.PolicyName) does not have any assignments!" -ForegroundColor Yellow
					$assignmentGroup = "Policy does not have any assignments!"
				} else {
					# There were assignments in Policy but we could not find which Azure AD group is causing policy to be applied
					Write-Host "Warning: Could not resolve Azure AD Group assignment for Policy $($ConfigurationPolicyReportState.PolicyName)!" -ForegroundColor Yellow
					
					$assignmentGroup = "unknown (possible user targeted group, nested group or removed assignment)"
				}

				$YodamiittiCustomGroupMembers = 'N/A'

				# Cast variable types to make sure column click based sorting works
				# Sorting may break if there are different kind of objects
				$properties = @{
					context                          = [String]$context
					odatatype                        = [String]$odatatype
					userPrincipalName                = [String]$ConfigurationPolicyReportState.UPN
					displayname                      = [String]$ConfigurationPolicyReportState.PolicyName
					assignmentIntent                 = [String]$assignmentIntent
					IncludeExclude                   = [String]$PolicyIncludeExclude
					assignmentGroup                  = [String]$assignmentGroup
					YodamiittiCustomGroupMembers     = [String]$YodamiittiCustomGroupMembers
					assignmentGroupId 				 = $null
					state                            = [String]$ConfigurationPolicyReportState.PolicyStatus
					YodamiittiCustomMembershipType   = [String]''
					id                               = $ConfigurationPolicyReportState.PolicyId
					filter							 = [String]''
					filterId						 = $null
					filterMode						 = [String]''
					filterTooltip					 = [String]''
					AssignmentGroupToolTip 			 = [String]''
					displayNameToolTip               = [String]''
				}

				$CustomObject = New-Object -TypeName PSObject -Prop $properties
				$script:ConfigurationsAssignmentsObservableCollection += $CustomObject
			}

			$lastDeviceConfigurationId = $ConfigurationPolicyReportState.PolicyId
		}
		
		# Filter out duplicate Policies
		# Intune shows applied policies to system (device) and possibly all users logged in to device
		# Combine same context/policy/state/assignmentGroup/Filter policies to one policy entry
		
		# DEBUG
		#$script:ConfigurationsAssignmentsObservableCollection | ConvertTo-Json -Depth 5 | Set-Clipboard

		# Get unique Policies eg. remove duplicates
		# Challenge is that -Unique selects first object from all duplicate objects
		# and that first object can have any value in userPrincipalName property
		$script:ConfigurationsAssignmentsObservableCollectionUnique = $script:ConfigurationsAssignmentsObservableCollection | Sort-Object -Property id,context,odatatype,displayName,IncludeExclude,state,assignmentGroup,filter,filterMode -Unique

		# Change PrimaryUser UPN to if found from assignments
		# Secondary change to device (which is empty value)
		foreach($PolicyInGrid in $script:ConfigurationsAssignmentsObservableCollectionUnique) {
			if(($PolicyInGrid.userPrincipalName -eq $PrimaryUserUserPrincipalName) -and ($primaryUserId -ne '00000000-0000-0000-0000-000000000000')) {
				# Policy UPN value is same than Intune device Primary User and PrimaryUser does exist
				
				# No change needed so continue to next policy in foreach loop
				Continue
			} elseif(($PolicyInGrid.userPrincipalName -eq $Script:LatestCheckedinUser.UserPrincipalName) -and ($primaryUserId -eq '00000000-0000-0000-0000-000000000000')) {
				# Policy UPN value is same than latest checked-in user and there is NO PrimaryUser
				
				# No change needed so continue to next policy in foreach loop
				Continue
			} else {
				# Policy UPN and Primary User values are different
				
				# Get duplicate policies from original list
				$DuplicatePolicyObjects = $script:ConfigurationsAssignmentsObservableCollection | Where-Object { ($_.id -eq $PolicyInGrid.id) -and ($_.context -eq $PolicyInGrid.context) -and ($_.odatatype -eq $PolicyInGrid.odatatype) -and ($_.displayName -eq $PolicyInGrid.displayName) -and ($_.IncludeExclude -eq $PolicyInGrid.IncludeExclude) -and ($_.state -eq $PolicyInGrid.state) -and ($_.assignmentGroup -eq $PolicyInGrid.assignmentGroup) -and ($_.filter -eq $PolicyInGrid.filter) -and ($_.filterMode -eq $PolicyInGrid.filterMode) }

				# Get userPrincipalNames in duplicate entries
				$UserPrincipalNames = $DuplicatePolicyObjects | Select-Object -ExpandProperty userPrincipalName
				
				# Check if primaryUser UPN was listed in duplicate policy entries
				if(($UserPrincipalNames -contains $PrimaryUserUserPrincipalName) -and ($primaryUserId -ne '00000000-0000-0000-0000-000000000000')) {
					$PolicyInGrid.userPrincipalName = $PrimaryUserUserPrincipalName
				} elseif(($UserPrincipalNames -contains $Script:LatestCheckedinUser.UserPrincipalName) -and ($primaryUserId -eq '00000000-0000-0000-0000-000000000000')) {
					$PolicyInGrid.userPrincipalName = $Script:LatestCheckedinUser.UserPrincipalName
				} else {
					# If primary user was not listed in duplicate policy entries
					# then we don't show any username for this specific policy
					$PolicyInGrid.userPrincipalName = ''
				}
			}
		}

		if($script:ConfigurationsAssignmentsObservableCollectionUnique.Count -gt 1) {
			# ItemsSource works if we are sorting 2 or more objects
			$WPFlistView_ConfigurationsAssignments.Itemssource = $script:ConfigurationsAssignmentsObservableCollectionUnique | Sort-Object displayName,userPrincipalName
		} else {
			# Only 1 object so we can't do sorting
			# If we try to sort here then our object array breaks and it does not work for ItemsSource
			# Cast as array because otherwise it will fail
			$WPFlistView_ConfigurationsAssignments.Itemssource = [array]$script:ConfigurationsAssignmentsObservableCollectionUnique
		}
	} else {
		Write-Host "Skipped Applications and Configurations Assignment report"
	}

	
    # Status textBox
    $DateTime = Get-Date -Format 'yyyy-MM-dd HH:mm.ss'
    $WPFbottom_textBox.Text = "Device details updated $DateTime"

	Write-Host "`nReport creation done"
	Write-Host "###############################################################################"
	Write-Host

	# Return information if this should be rerun with update cache data
	return $UnknownAssignmentGroupFound

}

Function Clear-UIData {

	Write-Verbose "Clear UI data"
	
	$WPFIntuneDeviceDetails_textBox_DeviceName.Text = ''
	
	# DeviceName ToolTip
	$WPFIntuneDeviceDetails_textBox_DeviceName_ToolTip_DeviceName.Text = ''
	$WPFIntuneDeviceDetails_textBox_DeviceName_ToolTip_DeviceProperties.Text = ''
	$WPFIntuneDeviceDetails_textBox_DeviceName_ToolTip_extensionAttributes.Text = ''
	
	$WPFManufacturer_textBox.Text = ''
	$WPFModel_textBox.Text = ''
	$WPFSerial_textBox.Text = ''
	$WPFWiFi_textBox.Text = ''
	$WPFOSVersion_textBox.Text = ''
	$WPFLanguage_textBox.Text = ''
	$WPFStorage_textBox.Text = ''
	$WPFEthernetMAC_textBox.Text = ''
	$WPFCompliance_textBox.Text = ''
	$WPFisEncrypted_textBox.Text = ''
	$WPFlastSync_textBox.Text = ''

	# PrimaryUser textBox and ToolTip
	$WPFprimaryUser_textBox.Text = ''
	$WPFtextBlock_primaryUser_textBox_ToolTip_UPN.Text = ''
	$WPFtextBlock_primaryUser_textBox_ToolTip_BasicInfo.Text = ''
	$WPFtextBlock_primaryUser_textBox_ToolTip_proxyAddresses.Text = ''
	$WPFtextBlock_primaryUser_textBox_ToolTip_otherMails.Text = ''
	$WPFtextBlock_primaryUser_textBox_ToolTip_onPremisesAttributes.Text = ''
	$WPFtextBlock_primaryUser_textBox_ToolTip_onPremisesExtensionAttributes.Text = ''
	
	$WPFAutopilotEnrolled_textBox.Text = ''
	$WPFAutopilotGroupTag_textBox.Text = ''
	$WPFAutopilotProfile_textBox.Text = ''

	# Recent checked-in user(s) textBox and ToolTip
	$WPFIntuneDeviceDetails_RecentCheckins_textBox.Text = ''
	$WPFtextBlock_LatestCheckInUser_textBox_ToolTip_UPN.Text = ''
	$WPFtextBlock_LatestCheckInUser_textBox_ToolTip_BasicInfo.Text = ''
	
	
	# Clear tooltips
	$WPFOSVersion_textBox.Tooltip = $null
	$WPFlastSync_textBox.Tooltip = $null
	$WPFAutopilotProfile_textBox.Tooltip = $null
	$WPFAutopilotGroupTag_textBox.Tooltip = $null
	

	$WPFTextBlock_TabItem_Device_GroupMemberships_Header_ToolTip.Text = ''
	$WPFTextBlock_TabItem_User_GroupMembershipsTAB_Header_ToolTip.Text = ''
	$WPFTextBlock_TabItem_LatestCheckedInUser_GroupMembershipsTAB_Header_ToolTip.Text = ''
	$WPFTabItem_LatestCheckedInUser_GroupMembershipsTAB_Header.Text = "Checked-In User Group Memberships"

	# Overview tabItem
	$WPFTextBlock_Overview_DeviceName.Text = 'DeviceName'
	$WPFTextBox_Overview_Device.Text = ''
	$WPFTextBlock_Overview_PrimaryUserName.Text = 'PrimaryUser'
	$WPFTextBox_Overview_PrimaryUser.Text = ''
	$WPFTextBlock_Overview_LatestCheckedInUserName.Text = 'Latest or Selected Checked-In User'
	$WPFTextBox_Overview_LatestCheckedInUser.Text = ''
	
	$WPFIntuneDeviceDetails_json_textBox.Text = ''
	$WPFAzureADDeviceDetails_json_textBox.Text = ''
	$WPFPrimaryUserDetails_json_textBox.Text = ''
	$WPFLatestCheckedInUser_json_textBox.Text = ''
	$WPFAutopilotDevice_json_textBox.Text = ''
	$WPFAutopilotEnrollmentProfile_json_textBox.Text = ''
	$WPFApplication_json_textBox.Text = ''
	$WPFConfiguration_json_textBox.Text = ''
	
	$WPFbottom_textBox.Text = ''

	# Clear Grid ListViews
	$WPFlistView_PrimaryUser_GroupMemberships.Itemssource = $null
	$WPFlistView_Device_GroupMemberships.Itemssource = $null
	$WPFlistView_LatestCheckedInUser_GroupMemberships.Itemssource = $null
	$WPFlistView_ApplicationAssignments.Itemssource = $null
	$WPFlistView_ConfigurationsAssignments.Itemssource = $null
	
	# Empty script wide variables
	$Script:IntuneManagedDevice = $null
	$Script:AutopilotDeviceWithAutpilotProfile = $null
	$Script:PrimaryUser = $null
	$Script:PrimaryUserGroupsMemberOf = $null
	$Script:AzureADDevice = $null
	$Script:LatestCheckedinUser = $null
	$Script:LatestCheckedInUserGroupsMemberOf = $null
	
	# Disable right click MenuItems
	$WPFIntuneDeviceDetails_textBox_DeviceName_Menu_OpenAutopilotDeviceInBrowser.isEnabled = $False
	$WPFtextBox_EnrollmentStatusPage_Menu_OpenESPProfileInBrowser.isEnabled = $False
	$WPFtextBox_EnrollmentRestrictions_Menu_OpenEnrollmentRestrictionProfilesInBrowser.isEnabled = $False
	$WPFtextBox_Compliance_Menu_OpenDeviceComplianceInBrowser.isEnabled = $False
	
	
	$WPFIntuneDeviceDetails_textBox_DeviceName_Menu_OpenDeviceInBrowser.isEnabled = $False
	$WPFIntuneDeviceDetails_textBox_DeviceName_Menu_OpenAzureADDeviceInBrowser.isEnabled = $False
	$WPFIntuneDeviceDetails_textBox_DeviceName_Menu_OpenAutopilotDeviceInBrowser.isEnabled = $False
	$WPFprimaryUser_textBox_Menu_OpenPrimaryUserInBrowser.isEnabled = $False
	$WPFAutopilotEnrolled_textBox_Menu_OpenAutopilotDeviceInBrowser.isEnabled = $False
	$WPFAutopilotProfile_textBox_Menu_OpenAutopilotDeploymentProfileInBrowser.isEnabled = $False
	
	$WPFListView_GridTabItem_Device_GroupMembershipsTAB_Menu_Copy_DynamicRules.isEnabled = $False
	$WPFListView_GridTabItem_Device_GroupMembershipsTAB_Menu_Copy_JSON.isEnabled = $False
	$WPFListView_GridTabItem_Device_GroupMembershipsTAB_Menu_Open_Group_In_Browser.isEnabled = $False
	
	$WPFListView_GridTabItem_PrimaryUser_GroupMembershipsTAB_Menu_Copy_DynamicRules.isEnabled = $False
	$WPFListView_GridTabItem_PrimaryUser_GroupMembershipsTAB_Menu_Copy_JSON.isEnabled = $False
	$WPFListView_GridTabItem_PrimaryUser_GroupMembershipsTAB_Menu_Open_Group_In_Browser.isEnabled = $False
	
	$WPFListView_GridTabItem_LatestCheckedInUser_GroupMembershipsTAB_Menu_Copy_DynamicRules.isEnabled = $False
	$WPFListView_GridTabItem_LatestCheckedInUser_GroupMembershipsTAB_Menu_Copy_JSON.isEnabled = $False
	$WPFListView_GridTabItem_LatestCheckedInUser_GroupMembershipsTAB_Menu_Open_Group_In_Browser.isEnabled = $False
	
	# Not used at the moment
	#$WPFlistView_ApplicationAssignments_Menu_Copy_ApplicationBasicInfo.isEnabled = $False
	$WPFlistView_ApplicationAssignments_Menu_Copy_JSON.isEnabled = $False
	$WPFlistView_ApplicationAssignments_Menu_Copy_DetectionRules_Powershell_to_Clipboard.isEnabled = $False
	$WPFlistView_ApplicationAssignments_Menu_Copy_requirementRules_Powershell_to_Clipboard.isEnabled = $False
	$WPFlistView_ApplicationAssignments_Menu_Open_Application_In_Browser.isEnabled = $False
	$WPFlistView_ApplicationAssignments_Menu_Open_ApplicationAssignmentBroup_In_Browser.isEnabled = $False
	$WPFlistView_ApplicationAssignments_Menu_Open_ApplicationAssignmentFilter_In_Browser.isEnabled = $False
	
	# Not used at the moment
	#$WPFlistView_ConfigurationsAssignments_Menu_Copy_ConfigurationBasicInfo.isEnabled = $False
	#$WPFlistView_ConfigurationsAssignments_Menu_Open_Configuration_In_Browser.isEnabled = $False
	
	$WPFlistView_ConfigurationsAssignments_Menu_Copy_JSON.isEnabled = $False
	$WPFlistView_ConfigurationsAssignments_Menu_Open_ConfigurationAssignmentBroup_In_Browser.isEnabled = $False
	$WPFlistView_ConfigurationsAssignments_Menu_Open_ConfigurationAssignmentFilter_In_Browser.isEnabled = $False
	
	$WPFprimaryUser_textBox_Menu_OpenPrimaryUserInBrowser.isEnabled = $False
	$WPFAutopilotEnrolled_textBox_Menu_OpenAutopilotDeviceInBrowser.isEnabled = $False
	$WPFAutopilotProfile_textBox_Menu_OpenAutopilotDeploymentProfileInBrowser.isEnabled = $False

	# Disable Latest Checked-In User right click menus
	$WPFLatest_CheckIn_User_Menu_Copy.isEnabled = $False
	$WPFLatest_CheckIn_User_Menu_Copy_Menu_OpenLatestCheckInUserInBrowser.isEnabled = $False

	
	# Clear Reload cache notification
	$WPFlabel_UnknownAssignmentGroupsFoundWarningText.Visibility = 'Collapsed'
}



##################################################################################################
#region Form specific functions closing etc...

$WPFButton_GridIntuneDeviceDetailsBorderTop_Search.Add_Click({
    
	Write-Verbose "Search button clicked"

	# Set Mouse cursor to Wait
	$Form.Cursor = [System.Windows.Input.Cursors]::Wait

	# Call search function
	Search-IntuneDevices

	# Set Mouse cursor to Arrow (default)
	$Form.Cursor = [System.Windows.Input.Cursors]::Arrow

})


$WPFButton_GridIntuneDeviceDetailsBorderTop_CreateReport.Add_Click({

	# Set this variable so any combobox selection change will not trigger
	# event. For example Select user -dropdown is one of these
	# which are configured and selected automatically during info gathering
	$script:IgnoreDropdownSelectionEvents = $True
    
	# Set Mouse cursor to wait
	$Form.Cursor = [System.Windows.Input.Cursors]::Wait
	
	Write-Verbose "Create report button clicked"

	# Clear UI Data
	Clear-UIData

	$IntuneDeviceObjectForReport = $WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.SelectedItem
	
	if($IntuneDeviceObjectForReport) {

		# Check if selectedItem is object type which has id property
		# This device id property can be used as is
		$IntuneDeviceId = $IntuneDeviceObjectForReport.id
		
		if($IntuneDeviceId) {
			# We got object because there is property id

			# If $IntuneDeviceId is guid then we pass it as is to Get-DeviceInformation function
			try {
				[System.Guid]::Parse($IntuneDeviceId) | Out-Null
					$IntuneDeviceIdIsValidGUID = $true
				} catch {
					$IntuneDeviceIdIsValidGUID = $false
				}

			if($IntuneDeviceIdIsValidGUID) {
				# $IntuneDeviceId is valid guid

				# Check if Reload cache is selected or not
				if($WPFcheckBox_ReloadCache.IsChecked) {
					# Get device information for report
					Write-Verbose "Reload cache checkBox enabled"
					$RerunWithCacheUpdate = Get-DeviceInformation $IntuneDeviceId -ReloadCacheData $true
				} else {
					# Update cache and
					# get device information for report
					$RerunWithCacheUpdate = Get-DeviceInformation $IntuneDeviceId -ReloadCacheData $false
				}
				
				# Uncheck Reload cache if we did run full report
				# Basic info was NOT selected
				if(($WPFcheckBox_ReloadCache.IsChecked) -and (-not $WPFcheckBox_SkipAppAndConfigurationAssignmentReport.IsChecked)) {
					$WPFcheckBox_ReloadCache.IsChecked = $False
				}
				
				if($RerunWithCacheUpdate) {
					$WPFlabel_UnknownAssignmentGroupsFoundWarningText.Visibility = 'Visible'
				} else {
					$WPFlabel_UnknownAssignmentGroupsFoundWarningText.Visibility = 'Collapsed'
				}
				
				# Set Overview-tab as selected tab on right side of window
				$WPFtabControlDetailsXAML.SelectedIndex = 0
				
			} else {
				# Search string is NOT valid guid
				# We should never get here
				Write-Verbose "IntuneDeviceId is NOT valid guid. This should not be possible"
				$WPFIntuneDeviceDetails_textBox_DeviceName.Text = "deviceId is not valid GUID"
				$WPFIntuneDeviceDetails_textBox_DeviceName.Foreground = "red"
				
				$script:IgnoreDropdownSelectionEvents = $False
				
				# Set Mouse cursor to Arrow (default)
				$Form.Cursor = [System.Windows.Input.Cursors]::Arrow
	
				return
			}
		} else {
			# We have something else
			# We should never get here
			Write-Verbose "Selected Intune Device for report is NOT object. Could not find deviceId (`$IntuneDeviceId: $IntuneDeviceId)We should never get here!"
			Write-Verbose "`$IntuneDeviceObjectForReport: $IntuneDeviceObjectForReport"
			$WPFIntuneDeviceDetails_textBox_DeviceName.Text = "Could not find deviceId"
			$WPFIntuneDeviceDetails_textBox_DeviceName.Foreground = "red"
		}
	} else {
		# We should never get here
		Write-Verbose "No device selected!"
		$WPFIntuneDeviceDetails_textBox_DeviceName.Text = "No device selected"
		$WPFIntuneDeviceDetails_textBox_DeviceName.Foreground = "red"
	}
	
	$script:IgnoreDropdownSelectionEvents = $False

	# Set Mouse cursor to Arrow (default)
	$Form.Cursor = [System.Windows.Input.Cursors]::Arrow

})


$WPFGridIntuneDeviceDetailsBorderTop_image_Search_X.Add_MouseLeftButtonDown( {

		$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.ItemsSource = $null
		$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.Items.Clear()
		
		# Update QuickFilters
		$Script:QuickSearchFilters = Update-QuickFilters

		$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.ItemsSource = $Script:QuickSearchFilters
		$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.DisplayMemberPath = 'QuickFilterName'
		$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.SelectedIndex = 0
		$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.IsDropDownOpen = $false
		
		# Set focus to devices ComboBox (or anywhere other than current ComboBox)
		# Reason is that we need to get GotFocus-event launched when clicking on Search ComboBox textbox
		# to clear values automatically. Se Focus needs to be somewhere else
		$WPFIntuneDeviceDetails_textBox_DeviceName.Focus()

		$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.IsEditable = $true
		
})


$WPFGridIntuneDeviceDetailsBorderTop_image_CreateReport_X.Add_MouseLeftButtonDown( {

		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.ItemsSource = $null
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.Items.Clear()
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.SelectedIndex = 0
		$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.IsDropDownOpen = $false

		$script:SelectedIntuneDeviceForReportTextboxSource = $null

		# Disable create report button
		$WPFButton_GridIntuneDeviceDetailsBorderTop_CreateReport.isEnabled = $False
		
		# Set button font to Normal
		$WPFButton_GridIntuneDeviceDetailsBorderTop_CreateReport.FontWeight = "Normal"
		
		# Clear Reload cache notification and checkBox
		$WPFlabel_UnknownAssignmentGroupsFoundWarningText.Visibility = 'Collapsed'
		$WPFcheckBox_ReloadCache.IsChecked = $False
		
		$WPFlabel_GridIntuneDeviceDetailsBorderTop_FoundXDevices.Visibility = 'Hidden'
		
})


#### Events ####

#$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.Add_SelectionChanged({
$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.Add_DropDownClosed( {
	$DropdownSelectedSearchObject = $WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.SelectedItem
	Write-Verbose "Selected item: $($DropdownSelectedSearchObject)"
	
	if($DropdownSelectedSearchObject.QuickFilterGraphAPIFilter -eq $null) {
		# Default text -> search text selected so we enable editing in ComboBox
		$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.IsEditable = $true
	} else {
		# Quick filter selected so we disable editing text on ComboBox
		$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.IsEditable = $false
	}
})


$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.Add_GotFocus( {
	#Write-Verbose "Search ComboBox GotFocus event launched"

	# DEBUG
	#Write-Verbose "Selected item $($WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.SelectedItem)"

	# Empty textBox text if default Search text is shown
	if(($WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.SelectedItem).QuickFilterName -eq "Search by deviceName, serialNumber, emailAddress, OS or id") {
		#Write-Verbose "Default search entry focused. Empty ComboBox textBox value"
		
		$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.Text = ''
	}
})


$WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.Add_DropDownClosed( {
	$DropdownSelectedSearchObject = $WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.SelectedItem
	#Write-Verbose "Selected item: $($DropdownSelectedSearchObject)"
	Write-Verbose "Selected item: $($DropdownSelectedSearchObject.deviceName)"

})

### DeviceName textbox Menu events

$WPFIntuneDeviceDetails_textBox_DeviceName_Menu_Copy.Add_Click({
	# Copy Device Name to clipboard
	$WPFIntuneDeviceDetails_textBox_DeviceName.Text | Set-Clipboard
})


$WPFIntuneDeviceDetails_textBox_DeviceName_Menu_OpenDeviceInBrowser.Add_Click({

	$IntuneDeviceId = $Script:IntuneManagedDevice.id
	if($IntuneDeviceId) {
		
		Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/overview/mdmDeviceId/$IntuneDeviceId"
	}
})


$WPFIntuneDeviceDetails_textBox_DeviceName_Menu_OpenAzureADDeviceInBrowser.Add_Click({

	$AzureADDeviceId = $Script:AzureADDevice.id
	if($AzureADDeviceId) {
		
		Start-Process "https://portal.azure.com/#blade/Microsoft_AAD_Devices/DeviceDetailsMenuBlade/Properties/objectId/$AzureADDeviceId"
	}
})

$WPFIntuneDeviceDetails_textBox_DeviceName_Menu_OpenAutopilotDeviceInBrowser.Add_Click({

	if($Script:IntuneManagedDevice.autopilotEnrolled) {
		
		$DeviceSerialNumber = $Script:IntuneManagedDevice.serialNumber
		
		# There isn't device specific page so we open main page and add device serial to clipboard
		# for fast search
		$DeviceSerialNumber | Set-Clipboard
		
		Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_Intune_Enrollment/AutoPilotDevicesBlade"
		
	}
})

### Device Compliance textBox Menu events
$WPFtextBox_Compliance_Menu_OpenDeviceComplianceInBrowser.Add_Click({

	$IntuneDeviceId = $Script:IntuneManagedDevice.id
	if($IntuneDeviceId) {
		
		Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/compliance/mdmDeviceId/$IntuneDeviceId"
	}
})


### PrimaryUser textbox Menu events
$WPFprimaryUser_textBox_Menu_Copy.Add_Click({
	# Copy Primary User email address to clipboard
	$WPFprimaryUser_textBox.Text | Set-Clipboard
})


$WPFprimaryUser_textBox_Menu_OpenPrimaryUserInBrowser.Add_Click({

	$PrimaryUserId = $Script:PrimaryUser.id
	
	if($PrimaryUserId) {
		
		Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_AAD_IAM/UserDetailsMenuBlade/Profile/userId/$PrimaryUserId"
	}
})


# Recent check-ins Menu events
$WPFLatest_CheckIn_User_Menu_Copy.Add_Click({
	# Copy Primary User email address to clipboard
	$Script:LatestCheckedinUser.userPrincipalName | Set-Clipboard
})


$WPFLatest_CheckIn_User_Menu_Copy_Menu_OpenLatestCheckInUserInBrowser.Add_Click({

	$UserId = $Script:LatestCheckedinUser.id
	
	if($UserId) {
		
		Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_AAD_IAM/UserDetailsMenuBlade/Profile/userId/$UserId"
	}
})


### Autopilot deviceEnrolled textBox menu event
$WPFAutopilotEnrolled_textBox_Menu_OpenAutopilotDeviceInBrowser.Add_Click({

	if($Script:IntuneManagedDevice.autopilotEnrolled) {
		
		$DeviceSerialNumber = $Script:IntuneManagedDevice.serialNumber
		
		# There isn't device specific page so we open main page and add device serial to clipboard
		# for fast search
		$DeviceSerialNumber | Set-Clipboard
		
		Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_Intune_Enrollment/AutoPilotDevicesBlade"
		
	}
})



# We need to download ESP Configuration Profiles to get right id	
# Disabled for now
$WPFtextBox_EnrollmentStatusPage_Menu_OpenESPProfileInBrowser.Add_Click({

	$EnrollmentStatusPageProfileId = $Script:EnrollmentStatusPagePolicyId
	if($EnrollmentStatusPageProfileId) {

		if($Script:EnrollmentStatusPageProfileName -eq 'Default ESP profile') {
			# Go to Enrollment Status Page overview page
			# Default Enrollment Status page has different ID than in report
			
			Start-Process 'https://endpoint.microsoft.com/#blade/Microsoft_Intune_Enrollment/EnrollmentStatusPageProfileListBlade'
			
		} else {
			# Go to Enrollment Status Page policy view
			Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_Intune_Enrollment/EnrollmentStatusPageMenuBlade/properties/profileId/$($EnrollmentStatusPageProfileId)_Windows10EnrollmentCompletionPageConfiguration"			
			
		}
	}
})


$WPFtextBox_EnrollmentRestrictions_Menu_OpenEnrollmentRestrictionProfilesInBrowser.Add_Click({

		Start-Process 'https://endpoint.microsoft.com/#blade/Microsoft_Intune_DeviceSettings/DevicesMenu/deviceTypeEnrollmentRestrictions'
})


### Autopilot Deployment profile textBox menu event
$WPFAutopilotProfile_textBox_Menu_OpenAutopilotDeploymentProfileInBrowser.Add_Click({

	$AutopilotIntendedDeploymentProfileId = $Script:AutopilotDeviceWithAutpilotProfile.intendedDeploymentProfile.id
	if($AutopilotIntendedDeploymentProfileId) {
		
		Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_Intune_Enrollment/AutopilotMenuBlade/properties/apProfileId/$AutopilotIntendedDeploymentProfileId"
	}
})




### Device Azure AD GroupMemberships
$WPFListView_GridTabItem_Device_GroupMembershipsTAB_Menu_Copy_DynamicRules.Add_Click( {

		# Cast as array so we can add multiple objects. Otherwise += will not work
		[array]$CopyInformationStringArray = $null

		foreach ($SelectedRowInGrid in $WPFlistView_Device_GroupMemberships.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			$GroupInformation = $SelectedRowInGrid | Select-Object -Property displayName,description,membershipRule | Format-List | Out-String
			$GroupInformation = $GroupInformation.Trim()
			
			$CopyInformationStringArray += $GroupInformation
			$CopyInformationStringArray += ''
		}

		$CopyInformationStringArray | Set-Clipboard
})

$WPFListView_GridTabItem_Device_GroupMembershipsTAB_Menu_Copy_JSON.Add_Click( {

		# Cast as array so we can add multiple objects. Otherwise += will not work
		[array]$CopyInformation = $null

		foreach ($SelectedRowInGrid in $WPFlistView_Device_GroupMemberships.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			$CopyInformation += $SelectedRowInGrid | Select-Object -Property * -ExcludeProperty YodamiittiCustomGroupType,YodamiittiCustomMembershipType,YodamiittiCustomGroupMembersCount

		}

		$CopyInformation | ConvertTo-Json | Set-Clipboard
})


$WPFListView_GridTabItem_Device_GroupMembershipsTAB_Menu_Open_Group_In_Browser.Add_Click({

		foreach ($SelectedRowInGrid in $WPFlistView_Device_GroupMemberships.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			$GroupId = $SelectedRowInGrid.id
			Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_AAD_IAM/GroupDetailsMenuBlade/Overview/groupId/$GroupId"
		}
})


### PrimaryUser GroupMemberships
$WPFListView_GridTabItem_PrimaryUser_GroupMembershipsTAB_Menu_Copy_DynamicRules.Add_Click( {

		# Cast as array so we can add multiple objects. Otherwise += will not work
		[array]$CopyInformationStringArray = $null

		foreach ($SelectedRowInGrid in $WPFlistView_PrimaryUser_GroupMemberships.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			$GroupInformation = $SelectedRowInGrid | Select-Object -Property displayName,description,membershipRule | Format-List | Out-String
			$GroupInformation = $GroupInformation.Trim()
			
			$CopyInformationStringArray += $GroupInformation
			$CopyInformationStringArray += ''
		}

		$CopyInformationStringArray | Set-Clipboard
})


$WPFListView_GridTabItem_PrimaryUser_GroupMembershipsTAB_Menu_Copy_JSON.Add_Click({

		# Cast as array so we can add multiple objects. Otherwise += will not work
		[array]$CopyInformation = $null

		foreach ($SelectedRowInGrid in $WPFlistView_PrimaryUser_GroupMemberships.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			$CopyInformation += $SelectedRowInGrid | Select-Object -Property * -ExcludeProperty YodamiittiCustom*

		}

		$CopyInformation | ConvertTo-Json | Set-Clipboard
})


$WPFListView_GridTabItem_PrimaryUser_GroupMembershipsTAB_Menu_Open_Group_In_Browser.Add_Click({

		foreach ($SelectedRowInGrid in $WPFlistView_PrimaryUser_GroupMemberships.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			$GroupId = $SelectedRowInGrid.id
			Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_AAD_IAM/GroupDetailsMenuBlade/Overview/groupId/$GroupId"
		}
})


### Latest Checked-In User GroupMemberships
$WPFListView_GridTabItem_LatestCheckedInUser_GroupMembershipsTAB_Menu_Copy_DynamicRules.Add_Click( {

		# Cast as array so we can add multiple objects. Otherwise += will not work
		[array]$CopyInformationStringArray = $null

		foreach ($SelectedRowInGrid in $WPFlistView_LatestCheckedInUser_GroupMemberships.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			$GroupInformation = $SelectedRowInGrid | Select-Object -Property displayName,description,membershipRule | Format-List | Out-String
			$GroupInformation = $GroupInformation.Trim()
			
			$CopyInformationStringArray += $GroupInformation
			$CopyInformationStringArray += ''
		}

		$CopyInformationStringArray | Set-Clipboard
})


$WPFListView_GridTabItem_LatestCheckedInUser_GroupMembershipsTAB_Menu_Copy_JSON.Add_Click({

		# Cast as array so we can add multiple objects. Otherwise += will not work
		[array]$CopyInformation = $null

		foreach ($SelectedRowInGrid in $WPFlistView_LatestCheckedInUser_GroupMemberships.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			$CopyInformation += $SelectedRowInGrid | Select-Object -Property * -ExcludeProperty YodamiittiCustom*

		}

		$CopyInformation | ConvertTo-Json | Set-Clipboard
})


$WPFListView_GridTabItem_LatestCheckedInUser_GroupMembershipsTAB_Menu_Open_Group_In_Browser.Add_Click({

		foreach ($SelectedRowInGrid in $WPFlistView_LatestCheckedInUser_GroupMemberships.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			$GroupId = $SelectedRowInGrid.id
			Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_AAD_IAM/GroupDetailsMenuBlade/Overview/groupId/$GroupId"
		}
})

### Application events

<#
$WPFlistView_ApplicationAssignments_Menu_Copy_ApplicationBasicInfo.Add_Click( {

		# Cast as array so we can add multiple objects. Otherwise += will not work
		[array]$CopyInformationStringArray = $null

		foreach ($SelectedRowInGrid in $WPFlistView_ApplicationAssignments.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			$GroupInformation = $SelectedRowInGrid | Select-Object -Property displayName,description | Format-List | Out-String
			$GroupInformation = $GroupInformation.Trim()
			
			$CopyInformationStringArray += $GroupInformation
			$CopyInformationStringArray += ''
		}

		$CopyInformationStringArray | Set-Clipboard
})
#>

$WPFlistView_ApplicationAssignments_Menu_Copy_JSON.Add_Click( {

		# Cast as array so we can add multiple objects. Otherwise += will not work
		[array]$CopyInformation = $null

		foreach ($SelectedRowInGrid in $WPFlistView_ApplicationAssignments.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"
			
			$CopyInformation += $Script:AppsWithAssignments | Where-Object -Property id -eq $SelectedRowInGrid.id

		}

		$CopyInformation | ConvertTo-Json -Depth 6 | Set-Clipboard
})

$WPFlistView_ApplicationAssignments_Menu_Copy_DetectionRules_Powershell_to_Clipboard.Add_Click( {

		# Cast as array so we can add multiple objects. Otherwise += will not work
		[array]$CopyInformation = $null

		foreach ($SelectedRowInGrid in $WPFlistView_ApplicationAssignments.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			# Only process win32 apps
			if($SelectedRowInGrid.odatatype -eq 'win32LobApp') {
				$App = $Script:AppsWithAssignments | Where-Object -Property id -eq $SelectedRowInGrid.id
				$AppDetectionRulePowershellScripts = $App.detectionRules | Where-Object '@odata.type' -eq '#microsoft.graph.win32LobAppPowerShellScriptDetection'

				foreach($DetectionRuleScript in $AppDetectionRulePowershellScripts) {
					$Base64EncodedScriptContent = $DetectionRuleScript.scriptContent
					if($Base64EncodedScriptContent) {
					
						# Decode Base64 content
						$b = [System.Convert]::FromBase64String("$Base64EncodedScriptContent")
						$DecodedScript = [System.Text.Encoding]::UTF8.GetString($b)
					}
					
					$CopyInformation += $DecodedScript
					$CopyInformation += "`n`n"
				}
			}
		}
		
		$CopyInformation | Set-Clipboard

})

$WPFlistView_ApplicationAssignments_Menu_Copy_requirementRules_Powershell_to_Clipboard.Add_Click( {

		# Cast as array so we can add multiple objects. Otherwise += will not work
		[array]$CopyInformation = $null

		foreach ($SelectedRowInGrid in $WPFlistView_ApplicationAssignments.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			# Only process win32 apps
			if($SelectedRowInGrid.odatatype -eq 'win32LobApp') {
				$App = $Script:AppsWithAssignments | Where-Object -Property id -eq $SelectedRowInGrid.id
				$AppRequirementRulesPowershellScripts = $App.requirementRules | Where-Object '@odata.type' -eq '#microsoft.graph.win32LobAppPowerShellScriptRequirement'

				foreach($RequirementRuleScript in $AppRequirementRulesPowershellScripts) {
					$ScriptDisplayName = $RequirementRuleScript.displayName
					$Base64EncodedScriptContent = $RequirementRuleScript.scriptContent
					if($Base64EncodedScriptContent) {
					
						# Decode Base64 content
						$b = [System.Convert]::FromBase64String("$Base64EncodedScriptContent")
						$DecodedScript = [System.Text.Encoding]::UTF8.GetString($b)
					}
					
					$CopyInformation += "#$ScriptDisplayName`n`n"
					$CopyInformation += $DecodedScript
					$CopyInformation += "`n`n"
				}
			}
		}

		$CopyInformation | Set-Clipboard
})


$WPFlistView_ApplicationAssignments_Menu_Open_Application_In_Browser.Add_Click({

		foreach ($SelectedRowInGrid in $WPFlistView_ApplicationAssignments.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"

			$AppId = $SelectedRowInGrid.id
			Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_Intune_Apps/SettingsMenu/2/appId/$AppId"
		}
})


$WPFlistView_ApplicationAssignments_Menu_Open_ApplicationAssignmentBroup_In_Browser.Add_Click({

		foreach ($SelectedRowInGrid in $WPFlistView_ApplicationAssignments.SelectedItems) {

			# DEBUG
			Write-Verbose "Selected row: $SelectedRowInGrid"

			$GroupId = $SelectedRowInGrid.assignmentGroupId
			
			if($GroupId) {
				Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_AAD_IAM/GroupDetailsMenuBlade/Overview/groupId/$GroupId"
			} else {
				Write-Verbose "AssignmentGroupId was empty"
			}
		}
})


$WPFlistView_ApplicationAssignments_Menu_Open_ApplicationAssignmentFilter_In_Browser.Add_Click({

		foreach ($SelectedRowInGrid in $WPFlistView_ApplicationAssignments.SelectedItems) {

			# DEBUG
			Write-Verbose "Selected row: $SelectedRowInGrid"

			$FilterId = $SelectedRowInGrid.filterId
			if($FilterId) {
				Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_Intune_DeviceSettings/AssignmentFilterSummaryBlade/assignmentFilterId/$FilterId"
			} else {
				Write-Verbose "FilterId was empty"
			}
		}
})

$WPFlistView_ApplicationAssignments.Add_SelectionChanged( {
		$SelectedApplication = $WPFlistView_ApplicationAssignments.SelectedItems
		$NumberOfApplicationsSelected = $SelectedApplication.Count

		#Write-Verbose "Selected Application $($SelectedApplication.displayName)"
		#Write-Verbose "Selected applications count is $NumberOfApplicationsSelected"
		
		# Continue only if 1 application is selected
		if($NumberOfApplicationsSelected -eq 1) {
			if($SelectedApplication.context -eq 'Device') {
				$WPFtabControl_Device_and_User_GroupMemberships.SelectedIndex = 0
				
				$WPFlistView_Device_GroupMemberships.SelectedItem = $Script:deviceGroupMemberships | Where-Object id -eq $SelectedApplication.assignmentGroupId
				
			}
			if($SelectedApplication.context -eq 'User') {
				$WPFtabControl_Device_and_User_GroupMemberships.SelectedIndex = 1
				
				$WPFlistView_PrimaryUser_GroupMemberships.SelectedItem = $Script:PrimaryUserGroupsMemberOf | Where-Object id -eq $SelectedApplication.assignmentGroupId
				
			}
			
			# Set Application JSON view text
			$IntuneApplicationData = $Script:AppsWithAssignments | Where-Object -Property id -eq $SelectedApplication.id
			$WPFApplication_json_textBox.Text = $IntuneApplicationData | ConvertTo-Json -Depth 6
			
			# MenuItem - "basic controls" always enabled
			$WPFlistView_ApplicationAssignments_Menu_Copy_JSON.isEnabled = $True
			$WPFlistView_ApplicationAssignments_Menu_Open_Application_In_Browser.isEnabled = $True
			
			# MenuItem - Enable Win32 Powershell scripts export only if there are Powershell script type requirement and/or detection check
			if($SelectedApplication.odatatype -eq 'Win32LobApp') {
				if($IntuneApplicationData.requirementRules.scriptContent) {
					# Win32 Application has requirement rule as Powershell script
					$WPFlistView_ApplicationAssignments_Menu_Copy_requirementRules_Powershell_to_Clipboard.isEnabled = $True
				} else {
					# Win32 Application does NOT have requirement rule as Powershell script
					$WPFlistView_ApplicationAssignments_Menu_Copy_requirementRules_Powershell_to_Clipboard.isEnabled = $False
				}

				if($IntuneApplicationData.detectionRules.scriptContent) {
					# Win32 Application has detection rule as Powershell script
					$WPFlistView_ApplicationAssignments_Menu_Copy_DetectionRules_Powershell_to_Clipboard.isEnabled = $True
				} else {
					# Win32 Application does NOT have detection rule as Powershell script
					$WPFlistView_ApplicationAssignments_Menu_Copy_DetectionRules_Powershell_to_Clipboard.isEnabled = $False
				}
			} else {
				# NOT Win32 Application so we disable these menuitems
				$WPFlistView_ApplicationAssignments_Menu_Copy_requirementRules_Powershell_to_Clipboard.isEnabled = $False
				$WPFlistView_ApplicationAssignments_Menu_Copy_DetectionRules_Powershell_to_Clipboard.isEnabled = $False
			}
			
			# MenuItem - Enable open Filter in browser if filter is specified
			# We cast value as String in script so we checking against '' and not against $null
			if($SelectedApplication.filterId -ne '') {
				$WPFlistView_ApplicationAssignments_Menu_Open_ApplicationAssignmentFilter_In_Browser.isEnabled = $True
			} else {
				$WPFlistView_ApplicationAssignments_Menu_Open_ApplicationAssignmentFilter_In_Browser.isEnabled = $False
			}
			
			# MenuItem - Open Assignment Group in Browser
			# Enable if there is AssignmentGroupId value
			# We cast value as String in script so we checking against '' and not against $null
			if($SelectedApplication.AssignmentGroupId -ne '') {
				$WPFlistView_ApplicationAssignments_Menu_Open_ApplicationAssignmentBroup_In_Browser.isEnabled = $True
			} else {
				$WPFlistView_ApplicationAssignments_Menu_Open_ApplicationAssignmentBroup_In_Browser.isEnabled = $False
			}
		}
	})


$WPFIntuneDeviceDetails_ApplicationAssignments_SelectUser_ComboBox.Add_SelectionChanged( {
		
		if(-not $script:IgnoreDropdownSelectionEvents) {
			
			# Set Mouse cursor to Wait
			# Does not work here
			#$Form.Cursor = [System.Windows.Input.Cursors]::Wait
			
			$SelectedUser = $WPFIntuneDeviceDetails_ApplicationAssignments_SelectUser_ComboBox.SelectedItem
			Write-Verbose "User selection change to user: $($SelectedUser.userPrincipalName)"
			
			$IntuneDeviceId = $Script:IntuneManagedDevice.id
			
			# Get Logged on users information
			Write-Host "Get group memberships for selected user $($SelectedUser.userPrincipalName)"
			Get-CheckedInUsersGroupMemberships -SelectedUser $SelectedUser
			
			Write-Host "Get Intune Apps and Intents for selected user $($SelectedUser.userPrincipalName)"
			Get-MobileAppIntentsForSpecifiedUser -UserId $SelectedUser.id -IntuneDeviceId $IntuneDeviceId
			Write-Host "Done"
		} else {
			Write-Verbose "Ignoring Select user -dropdown selection change"
		}

		# Set Mouse cursor to Arrow (default)
		# Does not work here
		#$Form.Cursor = [System.Windows.Input.Cursors]::Arrow

	})



### Configuration profiles listview events

$WPFlistView_ConfigurationsAssignments_Menu_Copy_JSON.Add_Click( {

		# Cast as array so we can add multiple objects. Otherwise += will not work
		[array]$CopyInformation = $null

		foreach ($SelectedRowInGrid in $WPFlistView_ConfigurationsAssignments.SelectedItems) {

			# DEBUG
			#Write-Verbose "Selected row: $SelectedRowInGrid"
			
			$CopyInformation += $Script:IntuneConfigurationProfilesWithAssignments | Where-Object -Property id -eq $SelectedRowInGrid.id

		}

		$CopyInformation | ConvertTo-Json -Depth 6 | Set-Clipboard
})


$WPFlistView_ConfigurationsAssignments_Menu_Open_ConfigurationAssignmentBroup_In_Browser.Add_Click({

		foreach ($SelectedRowInGrid in $WPFlistView_ConfigurationsAssignments.SelectedItems) {

			# DEBUG
			Write-Verbose "Selected row: $SelectedRowInGrid"

			$GroupId = $SelectedRowInGrid.assignmentGroupId
			
			if($GroupId) {
				Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_AAD_IAM/GroupDetailsMenuBlade/Overview/groupId/$GroupId"
			} else {
				Write-Verbose "AssignmentGroupId was empty"
			}
		}
})


$WPFlistView_ConfigurationsAssignments_Menu_Open_ConfigurationAssignmentFilter_In_Browser.Add_Click({

		foreach ($SelectedRowInGrid in $WPFlistView_ConfigurationsAssignments.SelectedItems) {

			# DEBUG
			Write-Verbose "Selected row: $SelectedRowInGrid"

			$FilterId = $SelectedRowInGrid.filterId
			if($FilterId) {
				Start-Process "https://endpoint.microsoft.com/#blade/Microsoft_Intune_DeviceSettings/AssignmentFilterSummaryBlade/assignmentFilterId/$FilterId"
			} else {
				Write-Verbose "FilterId was empty"
			}
		}
})

$WPFlistView_ConfigurationsAssignments.Add_SelectionChanged( {
		$SelectedConfiguration = $WPFlistView_ConfigurationsAssignments.SelectedItems
		$NumberOfConfigurationsSelected = $SelectedConfiguration.Count

		#Write-Verbose "Selected Configuration $($SelectedConfiguration.displayName)"
		#Write-Verbose "Selected Configurations count is $NumberOfConfigurationsSelected"
		
		# Continue only if 1 configuration is selected
		if($NumberOfConfigurationsSelected -eq 1) {
			if($SelectedConfiguration.context -eq 'Device') {
				$WPFtabControl_Device_and_User_GroupMemberships.SelectedIndex = 0
				
				$WPFlistView_Device_GroupMemberships.SelectedItem = $Script:deviceGroupMemberships | Where-Object id -eq $SelectedConfiguration.assignmentGroupId
				
			}
			if($SelectedConfiguration.context -eq 'User') {
				$WPFtabControl_Device_and_User_GroupMemberships.SelectedIndex = 1
				
				$WPFlistView_PrimaryUser_GroupMemberships.SelectedItem = $Script:PrimaryUserGroupsMemberOf | Where-Object id -eq $SelectedConfiguration.assignmentGroupId
				
			}
			
			# Set Configuration JSON view text
			$WPFConfiguration_json_textBox.Text = $Script:IntuneConfigurationProfilesWithAssignments | Where-Object -Property id -eq $SelectedConfiguration.id | ConvertTo-Json -Depth 6
			
			# MenuItem - "basic controls" always enabled
			$WPFlistView_ConfigurationsAssignments_Menu_Copy_JSON.isEnabled = $True
			
			# MenuItem - Enable open Filter in browser if filter is specified
			# We cast value as String in script so we checking against '' and not against $null
			if($SelectedConfiguration.filterId -ne '') {
				$WPFlistView_ConfigurationsAssignments_Menu_Open_ConfigurationAssignmentFilter_In_Browser.isEnabled = $True
			} else {
				$WPFlistView_ConfigurationsAssignments_Menu_Open_ConfigurationAssignmentFilter_In_Browser.isEnabled = $False
			}
			
			# MenuItem - Open Assignment Group in Browser
			# Enable if there is AssignmentGroupId value
			# We cast value as String in script so we checking against '' and not against $null
			if($SelectedConfiguration.AssignmentGroupId -ne '') {
				$WPFlistView_ConfigurationsAssignments_Menu_Open_ConfigurationAssignmentBroup_In_Browser.isEnabled = $True
			} else {
				$WPFlistView_ConfigurationsAssignments_Menu_Open_ConfigurationAssignmentBroup_In_Browser.isEnabled = $False
			}
		}
	})


$WPFlistView_Device_GroupMemberships.Add_SelectionChanged( {
		$SelectedAzureADGroup = $WPFlistView_Device_GroupMemberships.SelectedItems
		$NumberOfAzureADGroupsSelected = $SelectedAzureADGroup.Count

		#Write-Verbose "Selected AzureADGroup $($SelectedAzureADGroup)"
		#Write-Verbose "Selected AzureADGroup count is $NumberOfAzureADGroupsSelected"
		
		# Continue only if 1 configuration is selected
		if($NumberOfAzureADGroupsSelected -eq 1) {

			# Set Selected Azure AD Group JSON view text
			$WPFAzureAD_Group_json_textBox.Text = $SelectedAzureADGroup | Select-Object -Property * -ExcludeProperty YodamiittiCustom* | ConvertTo-Json -Depth 6
		}
	})

$WPFlistView_PrimaryUser_GroupMemberships.Add_SelectionChanged( {
		$SelectedAzureADGroup = $WPFlistView_PrimaryUser_GroupMemberships.SelectedItems
		$NumberOfAzureADGroupsSelected = $SelectedAzureADGroup.Count

		#Write-Verbose "Selected AzureADGroup $($SelectedAzureADGroup)"
		#Write-Verbose "Selected AzureADGroup count is $NumberOfAzureADGroupsSelected"
		
		# Continue only if 1 configuration is selected
		if($NumberOfAzureADGroupsSelected -eq 1) {

			# Set Selected Azure AD Group JSON view text
			$WPFAzureAD_Group_json_textBox.Text = $SelectedAzureADGroup | Select-Object -Property * -ExcludeProperty YodamiittiCustom* | ConvertTo-Json -Depth 6
		}
	})

$WPFlistView_LatestCheckedInUser_GroupMemberships.Add_SelectionChanged( {
		$SelectedAzureADGroup = $WPFlistView_LatestCheckedInUser_GroupMemberships.SelectedItems
		$NumberOfAzureADGroupsSelected = $SelectedAzureADGroup.Count

		#Write-Verbose "Selected AzureADGroup $($SelectedAzureADGroup)"
		#Write-Verbose "Selected AzureADGroup count is $NumberOfAzureADGroupsSelected"
		
		# Continue only if 1 configuration is selected
		if($NumberOfAzureADGroupsSelected -eq 1) {

			# Set Selected Azure AD Group JSON view text
			$WPFAzureAD_Group_json_textBox.Text = $SelectedAzureADGroup | Select-Object -Property * -ExcludeProperty YodamiittiCustom* | ConvertTo-Json -Depth 6
		}
	})


### AboutBottom events

$WPFAboutTAB_textBox_author_Menu_Copy.Add_Click({

	'Petri.Paavola@yodamiitti.fi' | Set-Clipboard 
})


$WPFAboutTAB_textBox_author_Menu_OpenEmailAddress.Add_Click({
	$Subject = "IntuneDeviceDetailsGUI.ps1 $ScriptVersion feedback"
	Start-Process "mailto:Petri.Paavola@yodamiitti.fi?subject=$Subject"
})


$WPFAboutTAB_textBox_github_link_Menu_Copy.Add_Click({
	'https://github.com/petripaavola/IntuneDeviceDetailsGUI' | Set-Clipboard 
})


$WPFAboutTAB_textBox_github_link_Menu_OpenDeviceInBrowser.Add_Click({
	Start-Process 'https://github.com/petripaavola/IntuneDeviceDetailsGUI'
})


$WPFAboutTAB_image_Yodamiitti_Menu_VisitYodamiittiWebsite.Add_Click({
	Start-Process 'http://www.yodamiitti.com'
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

	# Do not try to sort empty listView
	if($view) {
		# Not used because we will be confused if some column was sorted earlier
		# because this remembers that column last sort ordering
		#$sort = $view.SortDescriptions[0].Direction

		<#
			# Sorting broke because header name is seen differently
			# after changing listview column header to bold by adding <GridViewColumn.Header> element in XAML
			# originally we got HeaderName as string but after change we get object [System.Windows.Controls.GridViewColumnHeader]
			# so we needed to figure out how to get Header name out from this object
		#>

		# This works if we don't specify <GridViewColumn.Header> element in XAML code
		# This value is String
		# Example XAML: <GridViewColumn Width="350" Header="DisplayName" DisplayMemberBinding="{Binding 'displayName'}"/> 
		$columnHeader = $_.OriginalSource.Column.Header

		if($columnHeader -is [System.Windows.Controls.GridViewColumnHeader]) {
			# Column header is not string
			# Example XAML:
			<#
				<GridViewColumn Width="350">
					<GridViewColumn.Header>
						<GridViewColumnHeader FontWeight="Bold">DisplayName</GridViewColumnHeader>
					</GridViewColumn.Header>
					<GridViewColumn.CellTemplate>
						<DataTemplate>
							<TextBlock FontWeight="Bold" Text="{Binding 'displayName'}" ToolTip="{Binding Path=description}">
							</TextBlock>
						</DataTemplate>
					</GridViewColumn.CellTemplate>
				</GridViewColumn>
			#>

			#Write-Verbose "Column header is not string. Instead it is $($columnHeader.GetType())"
			
			# Use below commands to debug what properties we have in our Header object when debugging
			#$columnHeader | Out-String | Set-Clipboard
			#$columnHeader.GetType() | Out-String | Set-Clipboard
			#Write-Verbose "`$($columnHeader|Out-String): $($columnHeader|Out-String)"
			
			#Write-Verbose "`$($columnHeader.Content): $($columnHeader.Content)"
			#Write-Verbose "`$($columnHeader.Tag): $($columnHeader.Tag)"
			
			# Column Header name is in property Content
			if($columnHeader.Content) {
				$columnHeader = $columnHeader.Content
			}
		}

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

		Write-Verbose "$columnHeader column header clicked, doing $direction sorting to column/table."

		$view.SortDescriptions.Clear()
		$sortDescription = New-Object System.ComponentModel.SortDescription($columnHeader, $direction)
		$view.SortDescriptions.Add($sortDescription)

		# Save info we clicked this column.
		# If we click next time same column then we just reverse sort order
		$script:ApplicationGridLastColumnClicked = $columnHeader
	}
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

	# Do not try to sort empty listView
	if($view) {
		# Not used because we will be confused if some column was sorted earlier
		# because this remembers that column last sort ordering
		#$sort = $view.SortDescriptions[0].Direction

		<#
			# Sorting broke because header name is seen differently
			# after changing listview column header to bold by adding <GridViewColumn.Header> element in XAML
			# originally we got HeaderName as string but after change we get object [System.Windows.Controls.GridViewColumnHeader]
			# so we needed to figure out how to get Header name out from this object
		#>
		
		# This works if we don't specify <GridViewColumn.Header> element in XAML code
		# This value is String
		# Example XAML: <GridViewColumn Width="350" Header="DisplayName" DisplayMemberBinding="{Binding 'displayName'}"/> 
		$columnHeader = $_.OriginalSource.Column.Header
		
		if($columnHeader -is [System.Windows.Controls.GridViewColumnHeader]) {
			# Column header is not string
			# Example XAML:
			<#
				<GridViewColumn Width="350">
					<GridViewColumn.Header>
						<GridViewColumnHeader FontWeight="Bold">DisplayName</GridViewColumnHeader>
					</GridViewColumn.Header>
					<GridViewColumn.CellTemplate>
						<DataTemplate>
							<TextBlock FontWeight="Bold" Text="{Binding 'displayName'}" ToolTip="{Binding Path=description}">
							</TextBlock>
						</DataTemplate>
					</GridViewColumn.CellTemplate>
				</GridViewColumn>
			#>
			
			#Write-Verbose "Column header is not string. Instead it is $($columnHeader.GetType())"
			
			# Use below commands to debug what properties we have in our Header object when debugging
			#$columnHeader | Out-String | Set-Clipboard
			#$columnHeader.GetType() | Out-String | Set-Clipboard
			#Write-Verbose "`$($columnHeader|Out-String): $($columnHeader|Out-String)"
			
			#Write-Verbose "`$($columnHeader.Content): $($columnHeader.Content)"
			#Write-Verbose "`$($columnHeader.Tag): $($columnHeader.Tag)"
			
			# Column Header name is in property Content
			if($columnHeader.Content) {
				$columnHeader = $columnHeader.Content
			}
		}
		
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

		Write-Verbose "$columnHeader column header clicked, doing $direction sorting to column/table."

		$view.SortDescriptions.Clear()
		$sortDescription = New-Object System.ComponentModel.SortDescription($columnHeader, $direction)
		$view.SortDescriptions.Add($sortDescription)

		# Save info we clicked this column.
		# If we click next time same column then we just reverse sort order
		$script:ConfigurationProfilesGridLastColumnClicked = $columnHeader
	}		
}

# Add ConfigurationProfiles ListView Grid column select event handler for sorting when grid column is clicked
$WPFlistView_ConfigurationsAssignments.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $ConfigurationProfilesGridColumnClickEventCode)

#####


$TenantId = $null

# Howto convert images to base64 and back in XAML
# http://vcloud-lab.com/entries/powershell/powershell-gui-encode-decode-images

$logo = @'
iVBORw0KGgoAAAANSUhEUgAABE0AAAHyCAYAAAAeOzAmAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAD7xJREFUeNrs3NFx29gZgFEpowKcZ8ITl6ASpErWW4IqsLYCt2BXQqUDlaDMUu/qQAt5lcRxvn3xECQudc4MBm8a4MclSHxD6vz5+fkMAAAAgP/1NyMAAAAA+H+iCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQLoyAfXv/fvo87y5P4FRudrvHe1eUfZmmzcd598sJnMrX+bXxxdzP7uc53FjZ32a8HWmtzcd7Ku9Tb9Ei780LrmEG/Fw1r4eX+8Nnl2xI97//vvPezF6JJizh5Y3m6gTO451LyZ59OJHXxj/NnR9cDbbWLq0J780HWsP813aaNteDhJN31gTwb36eAwAALO0lRGxfv8UBMAzRBAAAOAThBBiOaAIAAByKcAIMRTQBAAAOSTgBhiGaAAAAhyacAEMQTQAAgGMQToDVE00AAIBjEU6AVRNNAACAYxJOgNUSTQAAgGMTToBVEk0AAIA1EE6A1RFNAACAtRBOgFURTQAAgDURToDVEE0AAIC1EU6AVRBNAACANRJOgKMTTQAAgLUSToCjEk0AAIA1E06AoxFNAACAtRNOgKMQTQAAgBEIJ8DBiSYAAMAohBPgoEQTAABgJMIJcDAXRgAAP+XD/IH91hj4ztO83RvDtwdaD7N/enjd3vz98nXb9zrbztvfB5vF/eu94q27fL2GsHqiCQD8/EPAJ2Pg+4eh3e7x+q0PYZo2V68Ps5ydfZ3XxK018S0wL3G/HPGh+2ZeE3fWxOblHnHlFsEI/DwHAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEC4MAIATtzDvN0NdLzv5u1yob99P29Pg107gFPyyzRtrozh7IMRMArRBICTtts9fpl3X0Y53tcP09uF/vzNPI87qwLgaD4aAYzFz3MAAAAAgmgCAAAAEEQTAAAAgCCaAAAAAAT/CBb+2naaNqYAAADwRvmmCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAOHCCAAA9uLdNG2ujOHs0gj+4x/WxJ9zMAJgVKIJAMB+vMSCrTHwnY+vG8v4dcBjvt7tHu/e+oWbps3LvfLKEmYEfp4DAACM5tfd7vGLMQBLE00AAICRCCbAwYgmAADAKAQT4KBEEwAAYASCCXBwogkAALB2gglwFKIJAACwZoIJcDSiCQAAsFaCCXBUogkAALBGgglwdKIJAACwNoIJsAqiCQAAsCaCCbAaogkAALAWggmwKqIJAACwBoIJsDqiCQAAcGyCCbBKogkAAHBMggmwWqIJAABwLIIJsGqiCQAAcAyCCbB6ogkAAHBoggkwBNEEAAA4JMEEGIZoAgAAHIpgAgxFNAEAAA5BMAGGI5oAAABLE0yAIZ0/Pz+bAnv1/v20nXdXJgEn67f5g++tMQAAa+P5ln3zTRMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAAhAsjgL90vds93o18AtO0uZp321O4GPO1OB/9HObrcTvvPnlpAQDAGHzTBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAADhwggAOGXTtLmdd59MYki/7XaPtwusie28u1rgeO/m470eaQ3Px3u+wPG+zHZr+Q5puDWM+/CPa3jerl029sk3TQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIFwYAQD8lPt5u1ng717O2+eFjvnm9bj3bWs58MNr43qwY15qDX+Zt68DzeFpwTncDTSHEe/DS3lwS+OtE00A4CcfLna7x70/BEzTZtGH2QGPmcHMa+xpsAfkJdfwv5Z4zQ24Jh5Gevge8T4MLMfPcwAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEAQTQAAAACCaAIAAAAQRBMAAACAIJoAAAAABNEEAAAAIIgmAAAAAEE0AQAAAAiiCQAAAEC4MAIWcH8i5/F0IudwZ0muxsOJXI8Hc1/0Xrfk63ap+9rdYGvtfrC/eyr3jjUbbQ0z7uenJ+Md8j7MG3b+/PxsCgAAAAA/8PMcAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAABBNAEAAAAIogkAAABAEE0AAAAAgmgCAAAAEEQTAAAAgCCaAAAAAATRBAAAACCIJgAAAADhDwEGADG6OdrIk7dVAAAAAElFTkSuQmCC
'@
$WPFbottomRightLogoimage.source = DecodeBase64Image -ImageBase64 $Logo

$picture_X = @'
iVBORw0KGgoAAAANSUhEUgAAABEAAAARCAIAAAC0D9CtAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAEnQAABJ0Ad5mH3gAAAE0SURBVDhPpZGxioNQEEX3B59aCBFFkNRpBcEihRb2qf0FK5FASAqLgNhaJYUIigi2FkHFnbyZVZM0u+ypfHPngNz5mv7OP5zH4+G6bpqm+FzTtq1pmlVV0Rudrussy2KMSZJ0vV4xQJqm2W63EKmqmuc5Dp9OkiSCIEAAiKIYxzFmdV0bhoFz4HA44Jz+7Xg8UsIY+OfzuSxLXddpxJjjOOM44vLSQRiGlHNNURR6MLbf74dhoL21AwRBQFsrbNvu+542OC8O4Ps+7XJ2ux1UStkPL87c0sxnk8DivLU0s24SIeetJc/zLpfLfABsEjeBp1MUhaZpGANzrafTiUZcg3twhTtw4M1mg9lbrVEU4RyAY+CQ/u1+v8uy/FkrgBqcgd7rDm6322etSJZl9MVZnN8yTd/Bx5q1HVOWAAAAAABJRU5ErkJggg==
'@
$WPFGridIntuneDeviceDetailsBorderTop_image_Search_X.source = DecodeBase64Image -ImageBase64 $picture_X
$WPFGridIntuneDeviceDetailsBorderTop_image_CreateReport_X.source = DecodeBase64Image -ImageBase64 $picture_X

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

# Set Reload cache Tooltip
$WPFcheckBox_ReloadCache.ToolTip = "Checking this will update Apps and Policies cache files.`n`nApplication and Configurations cache files are automatically reloaded every $Script:ReloadCacheEveryNDays days.`nApps and policies are also reloaded every time there has been change.`nSometimes Intune may not update lastModifiedDateTime property so data might not be up-to-date.`n`nUnknown assignment in report might be caused by removed assignment, nested Azure AD Groups or old data.`n`nUnknown assignments may also go away when devices sync with Intune and when Intune creates new monitoring report data."

# Set checkBox_SkipAppAndConfigurationAssignmentReport ToolTip
$WPFcheckBox_SkipAppAndConfigurationAssignmentReport.ToolTip = "Quicker search to get Basic info only.`n`nThis gets Intune Device and Primary User information only and skips Applications and Configurations Assignment report.`n`nCan be used to quickly check device and PrimaryUser information only."


##########################################################################################


####### Connect to Graph API
# Update Graph API schema to beta to get Win32LobApps and possible other new features also
Write-Host "Connecting to Intune using Powershell Intune-Module"
Update-MSGraphEnvironment -SchemaVersion 'beta'
$Success = $?

if (-not $Success) {
	Write-Host "Failed to update MSGraph Environment schema to Beta!`n" -ForegroundColor Red
	Write-Host "Make sure you have installed Intune Powershell module"
	Write-Host "You can install Intune module to your user account with command:`nInstall-Module -Name Microsoft.Graph.Intune -Scope CurrentUser" -ForegroundColor Yellow
	Write-Host "`nor you can install machine-wide Intune module with command:`nInstall-Module -Name Microsoft.Graph.Intune" -ForegroundColor Yellow
	Write-Host "More information: https://github.com/microsoft/Intune-PowerShell-SDK"
	Exit 1
}

$MSGraphEnvironment = Connect-MSGraph
$Success = $?

if ($Success -and $MSGraphEnvironment) {
	$TenantId = $MSGraphEnvironment.tenantId
	$AdminUserUPN = $MSGraphEnvironment.upn

	$WPFlabel_ConnectedAsUser_UserName.Content = $AdminUserUPN
	
} else {
	Write-Host "Could not connect to MSGraph!" -ForegroundColor Red
	Exit 1	
}


# Create tenant specific cache folder if not exist
# This is used for image caching used in GUI and Apps logos
$CacheFolderPath = "$PSScriptRoot\cache\$TenantId"
if(-not (Test-Path $CacheFolderPath)) {
	New-Item $CacheFolderPath -Itemtype Directory
	$Success = $?
	
	if($Success) {
		Write-Host "Created cache folder: $CacheFolderPath"
	} else {
		Write-Host "Could not create cache folder: $CacheFolderPath" -ForegroundColor Red
		Write-Host "Script will exit"
		Exit 1		
	}
}


# Set Search Quick Filters

$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.ItemsSource = $null
$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.Items.Clear()

# Update QuickFilters
$Script:QuickSearchFilters = Update-QuickFilters

$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.ItemsSource = $Script:QuickSearchFilters
$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.DisplayMemberPath = 'QuickFilterName'
$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.SelectedIndex = 0
$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.IsDropDownOpen = $false


# Run below if script parameter id was specified
if($id) {
	Write-Host "Getting Intune device information for Intune deviceId: $id"
	$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.Text = $id

	# Search for device to get device data to search and selected dropdownBoxes
	# Will recognize id (GUID) automatically
	Search-IntuneDevices

	$IntuneDeviceObjectForReport = $WPFComboBox_GridIntuneDeviceDetailsBorderTop_CreateReport.SelectedItem
	
	if($IntuneDeviceObjectForReport) {
		$IntuneDeviceId = $IntuneDeviceObjectForReport.id
		if($IntuneDeviceId) {
			# Get device information for report
			Get-DeviceInformation $IntuneDeviceId
		}
	}
}


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


# Do NOT Clear-UI


# Set Focus to search box
# Does not work, it does set focus to combobox but not to textBox inside combobox
#$WPFcomboBox_GridIntuneDeviceDetailsBorderTop_Search.Focus()


# Show Form
$Form.ShowDialog() | out-null

