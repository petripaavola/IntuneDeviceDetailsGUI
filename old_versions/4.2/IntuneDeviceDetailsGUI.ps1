<#
.SYNOPSIS
	Intune Device Details HTML report (Graph API)

.DESCRIPTION
	Console-based tool for searching Intune managed devices and generating modern HTML reports
	with complete device intelligence: apps, policies, scripts, assignments, group memberships,
	and conflict detection. Works in both Windows PowerShell 5.1 and PowerShell 7.x.

.PARAMETER Id
	Intune device ID (GUID) to generate report for. Accepts pipeline input from 
	Get-IntuneManagedDevice or other Graph API cmdlets. Alias: IntuneDeviceId

.PARAMETER SearchText
	Optional search text to pre-filter the device list in interactive mode.

.PARAMETER ReloadCache
	Forces reload of all cached data (apps, configuration profiles, scripts, assignments)
	from Microsoft Graph API, ignoring cache timestamps.

.PARAMETER SkipAssignments
	Skips downloading assignment information. Creates minimal reports faster but with incomplete data.

.PARAMETER DoNotOpenReportAutomatically
	Prevents the generated HTML report from opening automatically in the default browser.

.PARAMETER ExtendedReport
	Generates extended report including detailed policy settings, script contents, conflict detection,
	and full JSON data. Recommended for troubleshooting and documentation.

.PARAMETER OutputFolder
	Custom folder path for saving generated HTML reports. Defaults to 'reports' subfolder.

.EXAMPLE
	.\IntuneDeviceDetailsGUI.ps1
	
	Launches interactive mode with device search and report type selection.

.EXAMPLE
	.\IntuneDeviceDetailsGUI.ps1 -SearchText "DESKTOP"
	
	Pre-filters device list to show only devices matching "DESKTOP".

.EXAMPLE
	.\IntuneDeviceDetailsGUI.ps1 -Id 2e6e1d5f-b18a-44c6-989e-9bbb1efafbff -ExtendedReport
	
	Generates extended report directly for the specified device ID.

.NOTES
	Version: 4.2
	Author: Petri Paavola
	Requires: Microsoft.Graph.Authentication PowerShell module
	
	Cache files: cache\{TenantId}\ folder
	Reports: reports\ folder (or custom OutputFolder)

.LINK
	https://github.com/petripaavola/IntuneDeviceDetailsGUI
#>

[CmdletBinding(DefaultParameterSetName = 'interactive')]
param(
	[Parameter(Mandatory = $false,
			   ParameterSetName = 'id',
			   ValueFromPipeline = $true,
			   ValueFromPipelineByPropertyName = $true)]
	[ValidateScript({
		try {
			[System.Guid]::Parse($_) | Out-Null
			$true
		} catch {
			$false
		}
	})]
	[Alias('IntuneDeviceId')]
	[string]$Id,

	[Parameter(ParameterSetName = 'interactive')]
	[string]$SearchText,

	[switch]$ReloadCache,
	[switch]$SkipAssignments,
	[switch]$DoNotOpenReportAutomatically,
	[switch]$ExtendedReport,
	[string]$OutputFolder
)

$Version = '4.2'
$TimeOutBetweenGraphAPIRequests = 350
$GraphAPITop = 100
$script:ReloadCacheEveryNDays = 1
$script:ReportOutputFolder = if ($OutputFolder) { $OutputFolder } else { Join-Path -Path $PSScriptRoot -ChildPath 'reports' }
$script:ReportOutputFolder = [System.IO.Path]::GetFullPath($script:ReportOutputFolder)
$script:QuickSearchFilters = @()
$script:AllIntuneFilters = @()
$script:AppsWithAssignments = $null
$Script:IntuneConfigurationProfilesWithAssignments = @()
$script:GUIDHashtable = @{}

function Validate-GUID {
	param([string]$GUID)
	if (-not $GUID) { return $false }
	$pattern = '^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$'
	return [bool]($GUID -match $pattern)
}

function Add-GUIDToHashtable {
	param(
		[Parameter(Mandatory)][PSObject]$Object
	)
	
	# Extract ID from common property names
	$id = $null
	if ($Object.id) { $id = $Object.id }
	elseif ($Object.Id) { $id = $Object.Id }
	
	if (-not $id -or -not (Validate-GUID $id)) { return }
	
	if (-not $script:GUIDHashtable.ContainsKey($id)) {
		$value = @{
			Object = $Object
		}
		
		# Extract common name properties for quick access
		if ($Object.displayName) { $value.displayName = $Object.displayName }
		if ($Object.name) { $value.name = $Object.name }
		
		$script:GUIDHashtable[$id] = $value
	}
}

function Get-NameFromGUID {
	param(
		[Parameter(Mandatory)][string]$Id,
		[ValidateSet('displayName', 'name', 'any')]
		[string]$PreferredProperty = 'any'
	)
	if (-not $script:GUIDHashtable.ContainsKey($Id)) { return $null }
	$entry = $script:GUIDHashtable[$Id]
	
	if ($PreferredProperty -eq 'displayName' -and $entry.displayName) {
		return $entry.displayName
	}
	elseif ($PreferredProperty -eq 'name' -and $entry.name) {
		return $entry.name
	}
	else {
		# Return whichever is available
		if ($entry.displayName) { return $entry.displayName }
		if ($entry.name) { return $entry.name }
	}
	return $null
}

function Get-ObjectFromGUID {
	param(
		[Parameter(Mandatory)][string]$Id
	)
	if (-not $script:GUIDHashtable.ContainsKey($Id)) { return $null }
	return $script:GUIDHashtable[$Id].Object
}

function ConvertTo-LocalDateTimeString {
	param([Parameter(Mandatory = $false)][AllowNull()][object]$DateTimeValue)
	if (-not $DateTimeValue) { return 'n/a' }
	try {
		$parsed = [datetimeoffset]::Parse($DateTimeValue.ToString())
	} catch {
		return $DateTimeValue
	}
	return $parsed.LocalDateTime.ToString('yyyy-MM-dd HH:mm')
}

function Fix-UrlSpecialCharacters {
	param([Parameter(Mandatory)][string]$Url)
	$replacements = @(
		@(' ', '%20'),
		@('"', '%22'),
		@("'", '%27'),
		@('\\', '%5C'),
		@('@', '%40'),
		@('ä', '%C3%A4'),
		@('Ä', '%C3%84'),
		@('ö', '%C3%B6'),
		@('Ö', '%C3%96'),
		@('å', '%C3%A5'),
		@('Å', '%C3%85')
	)
	foreach ($pair in $replacements) {
		$Url = $Url.Replace($pair[0], $pair[1])
	}
	return $Url
}

function Get-CacheFolder {
	$tenantSegment = if ($script:TenantId) { $script:TenantId } else { 'default' }
	$cacheFolder = Join-Path -Path $PSScriptRoot -ChildPath "cache\$tenantSegment"
	Ensure-Directory -Path $cacheFolder | Out-Null
	return $cacheFolder
}

function Download-IntuneFilters {
	$url = 'https://graph.microsoft.com/beta/deviceManagement/assignmentFilters?`$select=*'
	$filters = Invoke-MGGraphGetRequestWithMSGraphAllPages $url

	# Add filter names to hashtable for easier access
	foreach ($filter in $filters) {
		Add-GUIDToHashtable -Object $filter
	}

	return [array]$filters
}

function Invoke-MGGraphPostRequest {
	param(
		[Parameter(Mandatory)][string]$Uri,
		[Parameter(Mandatory)][string]$Body
	)
	Start-Sleep -Milliseconds $TimeOutBetweenGraphAPIRequests
	$temporaryPath = Join-Path -Path $PSScriptRoot -ChildPath ("MgGraphRequest_{0}.json" -f (Get-Random))
	try {
		Invoke-MgGraphRequest -Uri $Uri -Method POST -Body $Body -OutputFilePath $temporaryPath -ContentType 'application/json' | Out-Null
		if (-not (Test-Path $temporaryPath)) { return $null }
		return Get-Content $temporaryPath -Raw | ConvertFrom-Json
	} finally {
		if (Test-Path $temporaryPath) {
			Remove-Item -Path $temporaryPath -ErrorAction SilentlyContinue
		}
	}
}


function Objectify_JSON_Schema_and_Data_To_PowershellObjects {
	param([Parameter(Mandatory)][psobject]$ReportData)

	# Objectify Intune configuration policies report json results to individual PowerShell objects

	if (-not $ReportData.Schema -or (-not $ReportData.Values)) { return @() }
	
	$rows = @()
	foreach ($row in $ReportData.Values) {
		$entry = [ordered]@{}
		for ($i = 0; $i -lt $ReportData.Schema.Count; $i++) {
			#$name = $ReportData.Schema[$i].Name
			$name = $ReportData.Schema[$i].Column
			$entry[$name] = $row[$i]
		}
		$rows += [pscustomobject]$entry
	}
	return $rows
}


function Download-IntunePostTypeReport {
	param(
		[Parameter(Mandatory)][string]$Uri,
		[Parameter(Mandatory)][string]$GraphAPIPostBody
	)

	$ConfigurationPoliciesReportForDevice = @()

	do {

		$GraphAPIPostBodyJSON = $GraphAPIPostBody | ConvertFrom-Json

		$top = $GraphAPIPostBodyJSON.top
		$skip = $GraphAPIPostBodyJSON.skip

		$response = Invoke-MGGraphPostRequest -Uri $Uri -Body $GraphAPIPostBody
		if ($response) {
			# Success

			if ($response.Schema -and $response.Values) {
				# Objectify report results
				$MgGraphRequestObjectified = Objectify_JSON_Schema_and_Data_To_PowershellObjects -ReportData $response

				# Save results to variable
				$ConfigurationPoliciesReportForDevice += $MgGraphRequestObjectified

				# Get Count of results
				$count = $MgGraphRequestObjectified.Count

				if($count -ge $top) {
					# Increase report skip-value with amount of results we got earlier (should be same as top)
					# to get next batch of results
					$skip += $count

					# Increase count in json and convert to text
					#$GraphAPIPostRequestJSON.top = $top
					$GraphAPIPostBodyJSON.skip = $skip

					# Convert json to text
					$GraphAPIPostBody = $GraphAPIPostBodyJSON | ConvertTo-Json -Depth 3

				} else {
					# Got all results
					Write-Verbose "Found $($ConfigurationPoliciesReportForDevice.Count) assignment objects"
				}
			}
		} else {
			Write-Verbose "Empty response from Graph API POST request to get report from $Uri"
			return $ConfigurationPoliciesReportForDevice
		}
	} while ($count -ge $top)

	return $ConfigurationPoliciesReportForDevice
}


function Download-IntuneConfigurationProfiles2 {
	param(
		[Parameter(Mandatory)][string]$GraphAPIUrl,
		[Parameter(Mandatory)][string]$jsonCacheFileName,
		[bool]$ReloadCacheData = $false
	)
	$cacheFolder = Get-CacheFolder
	$jsonCacheFilePath = Join-Path $cacheFolder $jsonCacheFileName
	
	if ((Test-Path $jsonCacheFilePath) -and (-not $ReloadCacheData)) {
		$fileDetails = Get-Item $jsonCacheFilePath
		$cacheAgeDays = (New-TimeSpan $fileDetails.LastWriteTimeUtc (Get-Date)).Days
		
		if ($cacheAgeDays -lt $script:ReloadCacheEveryNDays) {
			# Check if any configurations were modified after cache file timestamp

			# ORIGINAL
			# -UFormat type
			#$cacheFileLastWriteTimeUtc = Get-Date $fileDetails.LastWriteTimeUtc -UFormat '%Y-%m-%dT%H:%M:%S.000Z'

			# Real ISO 8601 format
			# You need to escape the colon characters in the format string to make sure every culture uses the correct format
			$cacheFileLastWriteTimeUtc = $fileDetails.LastWriteTimeUtc.ToString("yyyy-MM-ddTHH\:mm\:ss.fffffffZ")

			# Replace $select=* with a narrow select
			$GraphAPIUrlCheckUpdatesFix = $GraphAPIUrl -replace '\$select=\*', '$select=id,lastModifiedDateTime'

			# URL for checking for changes using lastModifiedDateTime filter
			$checkUrl = "$($GraphAPIUrlCheckUpdatesFix)&`$filter=lastModifiedDateTime%20gt%20$($cacheFileLastWriteTimeUtc)&`$orderby=lastModifiedDateTime%20desc&`$top=100"

			Write-Verbose "Checking for changes in $jsonCacheFileName since $cacheFileLastWriteTimeUtc"
			
			try {
				$changedConfigs = Invoke-MgGraphGetRequestWithMSGraphAllPages $checkUrl
				
				if (-not $changedConfigs) {
					# No changes found, use cache
					Write-Verbose "No changes detected, using cached $jsonCacheFileName"
					return Get-Content $jsonCacheFilePath -Raw | ConvertFrom-Json
				}
				Write-Host "Changes detected in $jsonCacheFileName, reloading from Graph API"
			} catch {
				# If delta query fails, fall back to full reload
				Write-Host "Delta query failed for $jsonCacheFileName (error: $_), performing full reload" -ForegroundColor Yellow
			}
		}
	}
	
	$data = Invoke-MgGraphGetRequestWithMSGraphAllPages $GraphAPIUrl
	if ($data) {
		$data | ConvertTo-Json -Depth 6 | Out-File $jsonCacheFilePath -Force
		return Get-Content $jsonCacheFilePath -Raw | ConvertFrom-Json
	}
	return @()
}


function Invoke-MGGraphGetRequestWithMSGraphAllPages {
	param([Parameter(Mandatory)][string]$url)
	$allGraphAPIData = @()
	do {
		Start-Sleep -Milliseconds $TimeOutBetweenGraphAPIRequests
		
		# Retry logic for transient Graph API failures
		$maxRetries = 5
		$retryCount = 0
		$response = $null
		
		while ($retryCount -lt $maxRetries) {
			try {
				$response = Invoke-MgGraphRequest -Uri $url -Method Get -OutputType PSObject -ContentType 'application/json'
				break  # Success - exit retry loop
			}
			catch {
				$retryCount++
				if ($retryCount -lt $maxRetries) {
					Write-Verbose "Graph API call failed (attempt $retryCount/$maxRetries). Retrying in 1 second... Error: $_"
					Start-Sleep -Seconds 1
				}
				else {
					Write-Warning "Graph API call failed after $maxRetries attempts: $_"
					return $null
				}
			}
		}

		if (-not $response) { return $null }

		if (Get-Member -InputObject $response -Name 'Value' -MemberType Properties) {
			$allGraphAPIData += $response.Value
			if (($response.'@odata.nextLink' -like 'https://*') -and (-not ($url.Contains('$top=')))) {
				$url = $response.'@odata.nextLink'
				continue
			}
			$url = $null
		} else {
			return $response
		}
	} while ($url)
	return $allGraphAPIData
}


function Add-AzureADGroupGroupTypeExtraProperties {
	param([array]$Groups)
	foreach ($group in $Groups) {
		if ($group.'@odata.type' -eq '#microsoft.graph.directoryRole') {
			$group | Add-Member -NotePropertyName 'YodamiittiCustomGroupType' -NotePropertyValue 'DirectoryRole' -Force
			$group | Add-Member -NotePropertyName 'YodamiittiCustomMembershipType' -NotePropertyValue 'Role' -Force
		} else {
			$membershipRule = $group.membershipRule
			if ([string]::IsNullOrEmpty($membershipRule)) {
				$group | Add-Member -NotePropertyName 'YodamiittiCustomGroupType' -NotePropertyValue 'Security' -Force
				$group | Add-Member -NotePropertyName 'YodamiittiCustomMembershipType' -NotePropertyValue 'Assigned' -Force
			} else {
				$group | Add-Member -NotePropertyName 'YodamiittiCustomGroupType' -NotePropertyValue 'Security' -Force
				$group | Add-Member -NotePropertyName 'YodamiittiCustomMembershipType' -NotePropertyValue 'Dynamic' -Force
			}
		}
	}
	return $Groups
}

function Add-AzureADGroupDevicesAndUserMemberCountExtraProperties {
	param([array]$Groups)
	
	if (-not $Groups -or $Groups.Count -eq 0) {
		return $Groups
	}
	
	Write-Verbose "Getting Entra ID groups member count for $($Groups.Count) groups"

	# Process groups in batches of 20 (Graph API batch limit)
	for ($i = 0; $i -lt $Groups.count; $i += 20) {

		# Create requests hashtables
		$requests_devices_count = @{ requests = @() }
		$requests_users_count = @{ requests = @() }

		# Create max 20 requests in for-loop
		for ($a = $i; (($a -lt $i + 20) -and ($a -lt $Groups.count)); $a += 1) {

			if ($Groups[$a].'@odata.type' -eq '#microsoft.graph.directoryRole') {
				# Azure DirectoryRole is not Entra ID Group
				$GraphAPIBatchEntry_DevicesCount = @{
					id     = ($a + 1).ToString()
					method = "GET"
					url    = "/directoryRoles/$($Groups[$a].id)"
				}
				$GraphAPIBatchEntry_UsersCount = @{
					id     = ($a + 1).ToString()
					method = "GET"
					url    = "/directoryRoles/$($Groups[$a].id)"
				}
			}
			else {
				# Entra ID Group - get transitive member counts
				$GraphAPIBatchEntry_DevicesCount = @{
					id     = ($a + 1).ToString()
					method = "GET"
					url    = "/groups/$($Groups[$a].id)/transitivemembers/microsoft.graph.device/`$count?ConsistencyLevel=eventual"
				}
				$GraphAPIBatchEntry_UsersCount = @{
					id     = ($a + 1).ToString()
					method = "GET"
					url    = "/groups/$($Groups[$a].id)/transitivemembers/microsoft.graph.user/`$count?ConsistencyLevel=eventual"
				}
			}

			$requests_devices_count.requests += $GraphAPIBatchEntry_DevicesCount
			$requests_users_count.requests += $GraphAPIBatchEntry_UsersCount
		}

		# Get device counts via batch API
		$requests_devices_count_JSON = $requests_devices_count | ConvertTo-Json -Depth 10
		$uri = 'https://graph.microsoft.com/beta/$batch'
		$AzureADGroups_Devices_MemberCount_Batch_Result = Invoke-MGGraphPostRequest -Uri $uri -Body $requests_devices_count_JSON.ToString()

		if ($AzureADGroups_Devices_MemberCount_Batch_Result) {
			# Process results for devices count batch requests
			foreach ($response in $AzureADGroups_Devices_MemberCount_Batch_Result.responses) {
				$GroupArrayIndex = $response.id - 1
				if ($response.status -eq 200) {
					if ($Groups[$GroupArrayIndex].'@odata.type' -eq '#microsoft.graph.directoryRole') {
						$Groups[$GroupArrayIndex] | Add-Member -MemberType NoteProperty -Name YodamiittiCustomGroupMembersCountDevices -Value 'N/A' -Force
					}
					else {
						$Groups[$GroupArrayIndex] | Add-Member -MemberType NoteProperty -Name YodamiittiCustomGroupMembersCountDevices -Value $response.body -Force
					}
				}
				else {
					Write-Warning "Error getting devices count for group $($Groups[$GroupArrayIndex].displayName)"
					$Groups[$GroupArrayIndex] | Add-Member -MemberType NoteProperty -Name YodamiittiCustomGroupMembersCountDevices -Value 'N/A' -Force
				}
			}
		}

		# Get user counts via batch API
		$requests_users_count_JSON = $requests_users_count | ConvertTo-Json -Depth 10
		$AzureADGroups_Users_MemberCount_Batch_Result = Invoke-MGGraphPostRequest -Uri $uri -Body $requests_users_count_JSON.ToString()

		if ($AzureADGroups_Users_MemberCount_Batch_Result) {
			# Process results for users count batch requests
			foreach ($response in $AzureADGroups_Users_MemberCount_Batch_Result.responses) {
				$GroupArrayIndex = $response.id - 1
				if ($response.status -eq 200) {
					if ($Groups[$GroupArrayIndex].'@odata.type' -eq '#microsoft.graph.directoryRole') {
						# Replace whole object with directoryRole details
						$Groups[$GroupArrayIndex] = $response.body
						$Groups[$GroupArrayIndex] | Add-Member -MemberType NoteProperty -Name YodamiittiCustomGroupMembersCountUsers -Value 'N/A' -Force
						$Groups[$GroupArrayIndex] | Add-Member -MemberType NoteProperty -Name YodamiittiCustomGroupMembersCountDevices -Value 'N/A' -Force
						$Groups[$GroupArrayIndex] | Add-Member -MemberType NoteProperty -Name YodamiittiCustomGroupType -Value 'DirectoryRole' -Force
					}
					else {
						$Groups[$GroupArrayIndex] | Add-Member -MemberType NoteProperty -Name YodamiittiCustomGroupMembersCountUsers -Value $response.body -Force
					}
				}
				else {
					Write-Warning "Error getting users count for group $($Groups[$GroupArrayIndex].displayName)"
					$Groups[$GroupArrayIndex] | Add-Member -MemberType NoteProperty -Name YodamiittiCustomGroupMembersCountUsers -Value 'N/A' -Force
				}
			}
		}
	}

	return $Groups
}

function Get-ApplicationsWithAssignments {
	param([bool]$ReloadCacheData = $false)

	$cacheFolder = Get-CacheFolder
	$cachePath = Join-Path $cacheFolder 'AllApplicationsWithAssignments.json'

	if ((Test-Path $cachePath) -and (-not $ReloadCacheData)) {
		$fileDetails = Get-Item $cachePath
		$ageDays = (New-TimeSpan $fileDetails.LastWriteTimeUtc (Get-Date)).Days
		
		if ($ageDays -lt $script:ReloadCacheEveryNDays) {
			# Check if any apps were modified after cache file timestamp
			$cacheFileLastWriteTimeUtc = Get-Date $fileDetails.LastWriteTimeUtc -UFormat '%Y-%m-%dT%H:%M:%S.000Z'
			Write-Host "Checking if apps were modified after $cacheFileLastWriteTimeUtc..." -ForegroundColor Cyan
			
			$checkUrl = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=lastModifiedDateTime%20gt%20$cacheFileLastWriteTimeUtc&`$top=1"
			$changedApps = Invoke-MGGraphGetRequestWithMSGraphAllPages $checkUrl
			
			if (-not $changedApps) {
				# No changes found, use cache
				Write-Host "No app changes detected. Using cached data." -ForegroundColor Green
				$cachedApps = Get-Content $cachePath -Raw | ConvertFrom-Json
				foreach ($app in $cachedApps) {
					Add-GUIDToHashtable -Object $app
				}
				return $cachedApps
			} else {
				Write-Host "App changes detected. Reloading all apps..." -ForegroundColor Yellow
			}
		}
	}

	Write-Host "Downloading all apps with assignments..." -ForegroundColor Cyan
	$url = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$expand=assignments&_=1577625591870'

	$apps = $null
	$apps = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
	if ($apps) {
		# Add App GUIDs to hashtable for easier access
		foreach ($app in $apps) {
			Add-GUIDToHashtable -Object $app
		}

		$apps | ConvertTo-Json -Depth 5 | Out-File $cachePath -Force
		return Get-Content $cachePath -Raw | ConvertFrom-Json
	}

	return @()
}

function Get-RemediationScriptsWithAssignments {
	param([bool]$ReloadCacheData = $false)

	# Note: Graph API does not support lastModifiedDateTime filtering for remediation scripts
	# Always download fresh data, but save to cache file for reference
	$cacheFolder = Get-CacheFolder
	$cachePath = Join-Path $cacheFolder 'RemediationScriptsAssignments.json'

	$url = 'https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts?$expand=assignments&select=id,displayName,description,createdDateTime,lastModifiedDateTime,runAsAccount,deviceHealthScriptType,assignments'

	$scripts = $null
	$scripts = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
	if ($scripts) {
		# Add Script GUIDs to hashtable for easier access
		foreach ($script in $scripts) {
			Add-GUIDToHashtable -Object $script
		}

		# Save to cache file for reference/debugging
		$scripts | ConvertTo-Json -Depth 5 | Out-File $cachePath -Force
		return $scripts
	}

	return @()
}

function Get-PlatformScriptsWithAssignments {
	param([bool]$ReloadCacheData = $false)

	$allScripts = @()

	# Windows platform scripts
	# Note: Graph API does not support lastModifiedDateTime filtering for platform scripts
	# Always download fresh data, but save to cache file for reference
	$cacheFolder = Get-CacheFolder
	$cachePath = Join-Path $cacheFolder 'WindowsPlatformScripts.json'

	$url = 'https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts?$expand=assignments'
	$scripts = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
	if ($scripts) {
		foreach ($script in $scripts) {
			$script | Add-Member -MemberType NoteProperty -Name 'ScriptPlatform' -Value 'Windows' -Force
			Add-GUIDToHashtable -Object $script
		}
		# Save to cache file for reference/debugging
		$scripts | ConvertTo-Json -Depth 5 | Out-File $cachePath -Force
		$allScripts += $scripts
	}

	# macOS shell scripts
	$cachePath = Join-Path $cacheFolder 'macOSShellScripts.json'

	$url = 'https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts?$expand=assignments'
	$scripts = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
	if ($scripts) {
		foreach ($script in $scripts) {
			$script | Add-Member -MemberType NoteProperty -Name 'ScriptPlatform' -Value 'macOS' -Force
			Add-GUIDToHashtable -Object $script
		}
		# Save to cache file for reference/debugging
		$scripts | ConvertTo-Json -Depth 5 | Out-File $cachePath -Force
		$allScripts += $scripts
	}

	# Linux bash scripts
	$cachePath = Join-Path $cacheFolder 'LinuxBashScripts.json'

	$url = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$expand=assignments&`$select=id,name,description,platforms,lastModifiedDateTime,technologies,settingCount,roleScopeTagIds,isAssigned,templateReference&`$filter=templateReference/TemplateFamily eq 'deviceConfigurationScripts'"
	$scripts = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
	if ($scripts) {
		foreach ($script in $scripts) {
			$script | Add-Member -MemberType NoteProperty -Name 'ScriptPlatform' -Value 'Linux' -Force
			Add-GUIDToHashtable -Object $script
		}
		# Save to cache file for reference/debugging
		$scripts | ConvertTo-Json -Depth 5 | Out-File $cachePath -Force
		$allScripts += $scripts
	}

	return $allScripts
}

function Get-PowerShellScriptContent {
	param (
		[Parameter(Mandatory = $true)]
		[string]$PowershellScriptPolicyId
	)

	try {
		Write-Verbose "Downloading script content for script ID: $PowershellScriptPolicyId"
		$url = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts/$PowershellScriptPolicyId"
		$scriptData = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
		
		if ($scriptData -and $scriptData.scriptContent) {
			# Decode base64 content
			$bytes = [System.Convert]::FromBase64String($scriptData.scriptContent)
			$scriptContentClearText = [System.Text.Encoding]::UTF8.GetString($bytes)
			return $scriptContentClearText
		}
		return $null
	}
	catch {
		Write-Warning "Failed to download script content for ID $PowershellScriptPolicyId : $_"
		return $null
	}
}

function Get-MacOSShellScriptContent {
	param (
		[Parameter(Mandatory = $true)]
		[string]$ShellScriptPolicyId
	)

	try {
		Write-Verbose "Downloading macOS shell script content for script ID: $ShellScriptPolicyId"
		$url = "https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts/$ShellScriptPolicyId"
		$scriptData = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
		
		if ($scriptData -and $scriptData.scriptContent) {
			# Decode base64 content
			$bytes = [System.Convert]::FromBase64String($scriptData.scriptContent)
			$scriptContentClearText = [System.Text.Encoding]::UTF8.GetString($bytes)
			return $scriptContentClearText
		}
		return $null
	}
	catch {
		Write-Warning "Failed to download macOS shell script content for ID $ShellScriptPolicyId : $_"
		return $null
	}
}

function Get-RemediationDetectionScriptContent {
	param (
		[Parameter(Mandatory = $true)]
		[string]$ScriptPolicyId
	)

	try {
		Write-Verbose "Downloading remediation detection script content for script ID: $ScriptPolicyId"
		$url = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/$ScriptPolicyId"
		$scriptData = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
		
		if ($scriptData -and $scriptData.detectionScriptContent) {
			# Decode base64 content
			$bytes = [System.Convert]::FromBase64String($scriptData.detectionScriptContent)
			$scriptContentClearText = [System.Text.Encoding]::UTF8.GetString($bytes)
			return $scriptContentClearText
		}
		return $null
	}
	catch {
		Write-Warning "Failed to download remediation detection script content for ID $ScriptPolicyId : $_"
		return $null
	}
}

function Get-RemediationRemediateScriptContent {
	param (
		[Parameter(Mandatory = $true)]
		[string]$ScriptPolicyId
	)

	try {
		Write-Verbose "Downloading remediation remediate script content for script ID: $ScriptPolicyId"
		$url = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/$ScriptPolicyId"
		$scriptData = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
		
		if ($scriptData -and $scriptData.remediationScriptContent) {
			# Decode base64 content
			$bytes = [System.Convert]::FromBase64String($scriptData.remediationScriptContent)
			$scriptContentClearText = [System.Text.Encoding]::UTF8.GetString($bytes)
			return $scriptContentClearText
		}
		return $null
	}
	catch {
		Write-Warning "Failed to download remediation remediate script content for ID $ScriptPolicyId : $_"
		return $null
	}
}

function Get-AppleEnrollmentProfileDetails {
	param (
		[Parameter(Mandatory = $true)]
		[string]$EnrollmentProfileName
	)

	try {
		Write-Verbose "Searching for Apple enrollment profile: $EnrollmentProfileName"
		
		# First, get all DEP onboarding settings with default profiles expanded
		$url = 'https://graph.microsoft.com/beta/deviceManagement/depOnboardingSettings?$expand=defaultiosenrollmentprofile,defaultmacosenrollmentprofile'
		$depSettings = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
		
		if (-not $depSettings) {
			Write-Verbose "No DEP onboarding settings found"
			return $null
		}
		
		# Phase 1: Check default enrollment profiles first
		foreach ($depToken in $depSettings) {
			# Check default iOS enrollment profile
			if ($depToken.defaultIosEnrollmentProfile -and $depToken.defaultIosEnrollmentProfile.displayName -eq $EnrollmentProfileName) {
				Write-Verbose "Found matching profile in default iOS enrollment profile of DEP token: $($depToken.tokenName)"
				return $depToken.defaultIosEnrollmentProfile
			}
			
			# Check default macOS enrollment profile
			if ($depToken.defaultMacOsEnrollmentProfile -and $depToken.defaultMacOsEnrollmentProfile.displayName -eq $EnrollmentProfileName) {
				Write-Verbose "Found matching profile in default macOS enrollment profile of DEP token: $($depToken.tokenName)"
				return $depToken.defaultMacOsEnrollmentProfile
			}
		}
		
		# Phase 2: If not found in defaults, search through all enrollment profiles for each DEP token
		Write-Verbose "Profile not found in default profiles, searching all enrollment profiles..."
		foreach ($depToken in $depSettings) {
			Write-Verbose "Searching enrollment profiles for DEP token: $($depToken.tokenName) (ID: $($depToken.id))"
			
			$profilesUrl = "https://graph.microsoft.com/beta/deviceManagement/depOnboardingSettings/$($depToken.id)/enrollmentProfiles"
			$enrollmentProfiles = Invoke-MGGraphGetRequestWithMSGraphAllPages $profilesUrl
			
			if ($enrollmentProfiles) {
				foreach ($profile in $enrollmentProfiles) {
					if ($profile.displayName -eq $EnrollmentProfileName) {
						Write-Verbose "Found matching profile: $($profile.displayName) in DEP token: $($depToken.tokenName)"
						return $profile
					}
				}
			}
		}
		
		Write-Verbose "Enrollment profile '$EnrollmentProfileName' not found in any DEP token"
		return $null
	}
	catch {
		Write-Warning "Failed to fetch Apple enrollment profile details for '$EnrollmentProfileName': $_"
		return $null
	}
}

function Get-SettingsCatalogPolicyDetails {
	param (
		[Parameter(Mandatory = $true)]
		[string]$PolicyId
	)

	try {
		Write-Verbose "Downloading Settings Catalog policy details for policy ID: $PolicyId"
		$url = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$PolicyId')/settings?`$expand=settingDefinitions"
		$settingsData = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
		return $settingsData
	}
	catch {
		Write-Warning "Failed to download Settings Catalog policy details for ID $PolicyId : $_"
		return $null
	}
}

function Analyze-SettingsCatalogConflicts {
	param (
		[Parameter(Mandatory = $true)]
		[array]$SettingsCatalogPolicies
	)

	if (-not $SettingsCatalogPolicies -or $SettingsCatalogPolicies.Count -lt 2) {
		# Need at least 2 policies to have conflicts
		return $null
	}

	Write-Verbose "Analyzing Settings Catalog conflicts across $($SettingsCatalogPolicies.Count) assigned policies..."
	
	$allSettings = @()
	
	# Extract all settings from all policies into flat structure
	foreach ($policy in $SettingsCatalogPolicies) {
		$policyId = $policy.id
		# Settings Catalog policies use 'name' property, not 'displayName'
		$policyName = if ($policy.name) { $policy.name } elseif ($policy.displayName) { $policy.displayName } else { $policyId }
		
		Write-Verbose "Processing policy: ID=$policyId, Name='$policyName'"
		
		# Get the downloaded settings details
		if (-not $script:GUIDHashtable.ContainsKey($policyId)) {
			Write-Verbose "Policy $policyId not found in GUIDHashtable"
			continue
		}
		
		$policyData = $script:GUIDHashtable[$policyId]
		if (-not $policyData.settingsRawData) {
			Write-Verbose "Policy $policyId has no settingsRawData"
			continue
		}
		
		Write-Verbose "Policy $policyId has $($policyData.settingsRawData.Count) settings"
		
		# Process each setting in the policy
		foreach ($settingItem in $policyData.settingsRawData) {
			$settingInstance = $settingItem.settingInstance
			$settingDefinitions = $settingItem.settingDefinitions
			
			# Extract setting info
			$extracted = Extract-SettingInfo -SettingInstance $settingInstance -SettingDefinitions $settingDefinitions -PolicyId $policyId -PolicyName $policyName
			if ($extracted) {
				$allSettings += $extracted
			}
		}
	}
	
	if ($allSettings.Count -eq 0) {
		return $null
	}
	
	Write-Verbose "Extracted $($allSettings.Count) settings from policies. Analyzing for conflicts..."
	
	# Group by both SettingDefinitionId AND SettingName to ensure we're comparing the exact same setting
	# This prevents false positives when different settings share the same parent category
	$grouped = $allSettings | Group-Object -Property SettingDefinitionId,SettingName | Where-Object { $_.Count -gt 1 }
	
	$conflicts = @()
	$warnings = @()
	
	# Known additive settings that should always be treated as warnings, not conflicts
	# These settings merge values across policies rather than conflicting
	$additiveSettingNames = @(
		'Excluded Extensions',
		'Excluded Paths', 
		'Excluded Processes',
		'Excluded File Extensions',
		'Excluded File Paths',
		'Excluded Process Names'
	)
	
	foreach ($group in $grouped) {
		$settingDef = $group.Name
		$instances = $group.Group
		
		# Sanity check: Filter out instances from the same policy (rare edge case)
		# Group by PolicyId to ensure we only report conflicts/warnings between different policies
		$uniquePolicyIds = $instances | Select-Object -ExpandProperty PolicyId -Unique
		if ($uniquePolicyIds.Count -lt 2) {
			Write-Verbose "Skipping setting '$($instances[0].SettingName)' - all instances are from the same policy (PolicyId: $($uniquePolicyIds[0]))"
			continue
		}
		
		# Check if this is a known additive setting
		$isAdditiveSetting = $false
		foreach ($instance in $instances) {
			if ($additiveSettingNames -contains $instance.SettingName) {
				$isAdditiveSetting = $true
				break
			}
		}
		
		# Get unique values
		$uniqueValues = $instances | Select-Object -ExpandProperty Value -Unique
		
		if ($isAdditiveSetting) {
			# Additive settings - always treat as warning even with different values
			$warnings += [PSCustomObject]@{
				SettingDefinitionId = $settingDef
				SettingName = $instances[0].SettingName
				Value = "Multiple values (additive)"
				Instances = $instances
				IsAdditive = $true
			}
		}
		elseif ($uniqueValues.Count -gt 1) {
			# CONFLICT - Different values for same setting
			$conflicts += [PSCustomObject]@{
				SettingDefinitionId = $settingDef
				SettingName = $instances[0].SettingName
				Instances = $instances
			}
		} else {
			# WARNING - Same setting configured in multiple policies with same value
			$warnings += [PSCustomObject]@{
				SettingDefinitionId = $settingDef
				SettingName = $instances[0].SettingName
				Value = $instances[0].Value
				Instances = $instances
				IsAdditive = $false
			}
		}
	}
	
	Write-Verbose "Found $($conflicts.Count) conflicts and $($warnings.Count) warnings"
	
	return [PSCustomObject]@{
		Conflicts = $conflicts
		Warnings = $warnings
		HasIssues = ($conflicts.Count -gt 0 -or $warnings.Count -gt 0)
	}
}

function Extract-SettingInfo {
	param (
		$SettingInstance,
		$SettingDefinitions,
		[string]$PolicyId,
		[string]$PolicyName,
		[string]$ParentPath = "",
		[switch]$IsInCollection
	)

	if (-not $SettingInstance) {
		return $null
	}

	$results = @()
	$settingDefId = $SettingInstance.settingDefinitionId
	$settingDef = $SettingDefinitions | Where-Object { $_.id -eq $settingDefId }
	
	if (-not $settingDef) {
		return $null
	}
	
	$displayName = $settingDef.displayName
	$currentPath = if ($ParentPath) { "$ParentPath > $displayName" } else { $displayName }
	
	# Extract value based on setting type
	switch ($SettingInstance.'@odata.type') {
		'#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance' {
			$choiceValue = $SettingInstance.choiceSettingValue.value
			$matchingOption = $settingDef.options | Where-Object { $_.itemId -eq $choiceValue }
			$valueDisplay = if ($matchingOption) { $matchingOption.displayName } else { $choiceValue }
			
			# Only add if not in a collection (to avoid false positives from collection items)
			if (-not $IsInCollection) {
				$results += [PSCustomObject]@{
					SettingDefinitionId = $settingDefId
					SettingName = $currentPath
					Value = $valueDisplay
					PolicyId = $PolicyId
					PolicyName = $PolicyName
				}
			}
			
			# Process children
			if ($SettingInstance.choiceSettingValue.children) {
				foreach ($child in $SettingInstance.choiceSettingValue.children) {
					$childResults = Extract-SettingInfo -SettingInstance $child -SettingDefinitions $SettingDefinitions -PolicyId $PolicyId -PolicyName $PolicyName -ParentPath $currentPath -IsInCollection:$IsInCollection
					if ($childResults) {
						$results += $childResults
					}
				}
			}
		}
		'#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance' {
			$simpleValue = $SettingInstance.simpleSettingValue.value
			
			# Only add if not in a collection
			if (-not $IsInCollection) {
				$results += [PSCustomObject]@{
					SettingDefinitionId = $settingDefId
					SettingName = $currentPath
					Value = $simpleValue
					PolicyId = $PolicyId
					PolicyName = $PolicyName
				}
			}
		}
		'#microsoft.graph.deviceManagementConfigurationGroupSettingInstance' {
			# Group settings - process children
			if ($SettingInstance.groupSettingValue.children) {
				foreach ($child in $SettingInstance.groupSettingValue.children) {
					$childResults = Extract-SettingInfo -SettingInstance $child -SettingDefinitions $SettingDefinitions -PolicyId $PolicyId -PolicyName $PolicyName -ParentPath $currentPath -IsInCollection:$IsInCollection
					if ($childResults) {
						$results += $childResults
					}
				}
			}
		}
		'#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance' {
			# Collection items are independent - don't compare across policies
			# Skip extracting values from collection items to avoid false positives
			Write-Verbose "Skipping collection setting: $currentPath (collections are additive, not conflicting)"
		}
		'#microsoft.graph.deviceManagementConfigurationChoiceSettingCollectionInstance' {
			# Multiple choice values - only compare if not already in a collection
			if (-not $IsInCollection -and $SettingInstance.choiceSettingCollectionValue) {
				$values = @()
				foreach ($choiceValue in $SettingInstance.choiceSettingCollectionValue) {
					$matchingOption = $settingDef.options | Where-Object { $_.itemId -eq $choiceValue.value }
					$values += if ($matchingOption) { $matchingOption.displayName } else { $choiceValue.value }
				}
				
				$results += [PSCustomObject]@{
					SettingDefinitionId = $settingDefId
					SettingName = $currentPath
					Value = ($values -join ', ')
					PolicyId = $PolicyId
					PolicyName = $PolicyName
				}
			}
		}
		'#microsoft.graph.deviceManagementConfigurationSimpleSettingCollectionInstance' {
			# Simple value collection - only compare if not already in a collection
			if (-not $IsInCollection -and $SettingInstance.simpleSettingCollectionValue) {
				$values = $SettingInstance.simpleSettingCollectionValue | ForEach-Object { $_.value }
				
				$results += [PSCustomObject]@{
					SettingDefinitionId = $settingDefId
					SettingName = $currentPath
					Value = ($values -join ', ')
					PolicyId = $PolicyId
					PolicyName = $PolicyName
				}
			}
		}
	}
	
	return $results
}

function Get-OmaSettingPlainTextValue {
	param (
		[Parameter(Mandatory = $true)]
		[string]$PolicyId,
		[Parameter(Mandatory = $true)]
		[string]$SecretReferenceValueId
	)

	try {
		Write-Verbose "Fetching encrypted OMA setting value for policy ID: $PolicyId, secret: $SecretReferenceValueId"
		$url = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations('$PolicyId')/getOmaSettingPlainTextValue(secretReferenceValueId='$SecretReferenceValueId')"
		$response = Invoke-MgGraphRequest -Method GET -Uri $url
		return $response.value
	}
	catch {
		Write-Warning "Failed to fetch encrypted OMA setting value for policy ID $PolicyId : $_"
		return $null
	}
}

function ConvertTo-ReadableSettingsCatalog {
	param (
		[Parameter(Mandatory = $true)]
		$SettingsData
	)

	if (-not $SettingsData -or $SettingsData.Count -eq 0) {
		return $null
	}

	$readableSettings = @()

	foreach ($settingItem in $SettingsData) {
		$settingInstance = $settingItem.settingInstance
		$settingDefinitions = $settingItem.settingDefinitions

		# Parse the setting recursively
		$parsedSetting = Parse-SettingInstance -SettingInstance $settingInstance -SettingDefinitions $settingDefinitions -IndentLevel 0
		if ($parsedSetting) {
			$readableSettings += $parsedSetting
		}
	}

	return ($readableSettings -join "`n`n")
}

function Parse-SettingInstance {
	param (
		$SettingInstance,
		$SettingDefinitions,
		[int]$IndentLevel = 0
	)

	if (-not $SettingInstance) {
		return $null
	}

	$indent = '  ' * $IndentLevel
	$output = @()

	# Find the setting definition for this instance
	$settingDefId = $SettingInstance.settingDefinitionId
	$settingDef = $SettingDefinitions | Where-Object { $_.id -eq $settingDefId }

	if ($settingDef) {
		$displayName = $settingDef.displayName
		
		# Determine the configured value based on setting type
		switch ($SettingInstance.'@odata.type') {
			'#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance' {
				# This is a choice setting (like a dropdown or toggle)
				$choiceValue = $SettingInstance.choiceSettingValue.value
				
				# Find the matching option in the definition
				$matchingOption = $settingDef.options | Where-Object { $_.itemId -eq $choiceValue }
				if ($matchingOption) {
					$configuredValue = $matchingOption.displayName
				} else {
					$configuredValue = $choiceValue
				}
				
				$output += "$indent$displayName : $configuredValue"
				
				# Process child settings if they exist
				if ($SettingInstance.choiceSettingValue.children -and $SettingInstance.choiceSettingValue.children.Count -gt 0) {
					foreach ($child in $SettingInstance.choiceSettingValue.children) {
						$childOutput = Parse-SettingInstance -SettingInstance $child -SettingDefinitions $SettingDefinitions -IndentLevel ($IndentLevel + 1)
						if ($childOutput) {
							$output += $childOutput
						}
					}
				}
			}
			'#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance' {
				# Simple setting (text, number, etc.)
				$simpleValue = $SettingInstance.simpleSettingValue.value
				$output += "$indent$displayName : $simpleValue"
			}
			'#microsoft.graph.deviceManagementConfigurationGroupSettingInstance' {
				# Group setting - has multiple child settings
				$output += "$indent$displayName"
				
				if ($SettingInstance.groupSettingValue.children -and $SettingInstance.groupSettingValue.children.Count -gt 0) {
					foreach ($child in $SettingInstance.groupSettingValue.children) {
						$childOutput = Parse-SettingInstance -SettingInstance $child -SettingDefinitions $SettingDefinitions -IndentLevel ($IndentLevel + 1)
						if ($childOutput) {
							$output += $childOutput
						}
					}
				}
			}
			'#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance' {
				# Group setting collection - has an array of group setting values
				# Each group setting value has children
				# Example: Firewall rules, Windows Hello for Business policies
				if ($SettingInstance.groupSettingCollectionValue -and $SettingInstance.groupSettingCollectionValue.Count -gt 0) {
					$groupIndex = 0
					foreach ($groupValue in $SettingInstance.groupSettingCollectionValue) {
						if ($groupValue.children -and $groupValue.children.Count -gt 0) {
							foreach ($child in $groupValue.children) {
								$childOutput = Parse-SettingInstance -SettingInstance $child -SettingDefinitions $SettingDefinitions -IndentLevel $IndentLevel
								if ($childOutput) {
									$output += $childOutput
								}
							}
							# Add blank line between collection items (e.g., between firewall rules)
							if ($groupIndex -lt ($SettingInstance.groupSettingCollectionValue.Count - 1)) {
								$output += ""
							}
						}
						$groupIndex++
					}
				}
			}
			'#microsoft.graph.deviceManagementConfigurationChoiceSettingCollectionInstance' {
				# Choice setting collection - has an array of choice values
				# Example: Interface Types, Network Types (Profiles) in firewall rules
				if ($SettingInstance.choiceSettingCollectionValue -and $SettingInstance.choiceSettingCollectionValue.Count -gt 0) {
					$values = @()
					foreach ($choiceValue in $SettingInstance.choiceSettingCollectionValue) {
						# Find the matching option in the definition
						$matchingOption = $settingDef.options | Where-Object { $_.itemId -eq $choiceValue.value }
						if ($matchingOption) {
							$values += $matchingOption.displayName
						} else {
							$values += $choiceValue.value
						}
					}
					$output += "$indent$displayName : $($values -join ', ')"
				}
			}
			'#microsoft.graph.deviceManagementConfigurationSimpleSettingCollectionInstance' {
				# Simple setting collection - has an array of simple values
				# Example: Reusable groups (Remote Address Dynamic Keywords) in firewall rules
				if ($SettingInstance.simpleSettingCollectionValue -and $SettingInstance.simpleSettingCollectionValue.Count -gt 0) {
					$values = @()
					foreach ($simpleValue in $SettingInstance.simpleSettingCollectionValue) {
						# Handle reference settings (GUID references)
						if ($simpleValue.'@odata.type' -eq '#microsoft.graph.deviceManagementConfigurationReferenceSettingValue') {
							# This is a reference to another setting (like a reusable group)
							# For now, display the GUID - could potentially resolve to name later
							$values += $simpleValue.value
						} else {
							# Regular simple value
							$values += $simpleValue.value
						}
					}
					$output += "$indent$displayName : $($values -join ', ')"
				}
			}
			Default {
				# Unknown type - try to extract basic info
				$output += "$indent$displayName : (unsupported setting type: $($SettingInstance.'@odata.type'))"
			}
		}
	}

	return ($output -join "`n")
}

function ConvertTo-ReadableOmaSettings {
	param($OmaSettings)

	if (-not $OmaSettings -or $OmaSettings.Count -eq 0) {
		return $null
	}

	$output = @()
	foreach ($setting in $OmaSettings) {
		$lines = @()
		
		# Display Name
		if ($setting.displayName) {
			$lines += "Name : $($setting.displayName)"
		}
		
		# Description
		if ($setting.description) {
			$lines += "Description : $($setting.description)"
		}
		
		# OMA-URI
		if ($setting.omaUri) {
			$lines += "OMA-URI : $($setting.omaUri)"
		}
		
		# Data type based on @odata.type
		$dataType = switch ($setting.'@odata.type') {
			'#microsoft.graph.omaSettingBoolean' { 'Boolean' }
			'#microsoft.graph.omaSettingString' { 'String' }
			'#microsoft.graph.omaSettingInteger' { 'Integer' }
			'#microsoft.graph.omaSettingStringXml' { 'String (XML file)' }
			'#microsoft.graph.omaSettingBase64' { 'Base64' }
			Default { $setting.'@odata.type' -replace '#microsoft\.graph\.', '' }
		}
		$lines += "Data type : $dataType"
		
		# Value - handle different types
		if ($null -ne $setting.value) {
			$valueDisplay = switch ($setting.'@odata.type') {
				'#microsoft.graph.omaSettingBoolean' { $setting.value.ToString() }
				'#microsoft.graph.omaSettingStringXml' {
					# For XML, show first 100 chars or indicate it's XML content
					if ($setting.fileName) {
						"$($setting.fileName) (XML content)"
					} else {
						$xmlPreview = $setting.value.ToString()
						if ($xmlPreview.Length -gt 100) {
							"$($xmlPreview.Substring(0, 100))..."
						} else {
							$xmlPreview
						}
					}
				}
				'#microsoft.graph.omaSettingBase64' {
					# For Base64, show file name if available
					if ($setting.fileName) {
						"$($setting.fileName) (Base64 encoded)"
					} else {
						"Base64 encoded data"
					}
				}
				Default { $setting.value.ToString() }
			}
			$lines += "Value : $valueDisplay"
		}
		
		$output += ($lines -join "`n")
	}

	return ($output -join "`n`n")
}

function ConvertTo-ReadableWin32LobApp {
	param(
		$AppData,
		[switch]$ExtendedReport
	)

	if (-not $AppData) {
		return $null
	}
	
	$odataType = $AppData.'@odata.type'
	if ($odataType -ne '#microsoft.graph.win32LobApp' -and $odataType -ne '#microsoft.graph.win32CatalogApp') {
		return $null
	}

	$output = @()
	
	# App Information
	if ($AppData.description) {
		$output += "Description : $($AppData.description)"
	}
	if ($AppData.publisher) {
		$output += "Publisher : $($AppData.publisher)"
	}
	if ($AppData.displayVersion) {
		$output += "App Version : $($AppData.displayVersion)"
	}
	
	# Program section
	$output += "`n--- Program ---"
	
	# Check for new install/uninstall script feature (January 2026)
	# Note: The batch apps query doesn't return activeInstallScript/activeUninstallScript properties
	# We use a dirty workaround: if we see placeholder command lines, fetch individual app to check for scripts
	$hasInstallScript = $AppData.activeInstallScript -and $AppData.activeInstallScript.targetId
	$hasUninstallScript = $AppData.activeUninstallScript -and $AppData.activeUninstallScript.targetId
	
	# Dirty workaround: Check for Microsoft's internal placeholder command lines that indicate scripts are used
	$suspectInstallPlaceholder = $AppData.installCommandLine -eq 'foobar.cmd'
	$suspectUninstallPlaceholder = $AppData.uninstallCommandLine -eq 'uninstall-foobar.cmd'
	
	# If both script properties are null AND we see the placeholder commands, fetch individual app
	if (($null -eq $AppData.activeInstallScript) -and ($null -eq $AppData.activeUninstallScript) -and 
	    ($suspectInstallPlaceholder -or $suspectUninstallPlaceholder) -and $AppData.id) {
		Write-Verbose "Win32App '$($AppData.displayName)': Detected placeholder commands, fetching individual app to check for scripts..."
		try {
			$individualAppUrl = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($AppData.id)?`$expand=assignments"
			$individualApp = Invoke-MGGraphGetRequestWithMSGraphAllPages $individualAppUrl
			
			if ($individualApp) {
				$hasInstallScript = $individualApp.activeInstallScript -and $individualApp.activeInstallScript.targetId
				$hasUninstallScript = $individualApp.activeUninstallScript -and $individualApp.activeUninstallScript.targetId
				Write-Verbose "  After individual fetch: hasInstallScript=$hasInstallScript, hasUninstallScript=$hasUninstallScript"
				
				# Update AppData with the fetched script info for later use
				if ($individualApp.activeInstallScript) {
					$AppData | Add-Member -NotePropertyName 'activeInstallScript' -NotePropertyValue $individualApp.activeInstallScript -Force
				}
				if ($individualApp.activeUninstallScript) {
					$AppData | Add-Member -NotePropertyName 'activeUninstallScript' -NotePropertyValue $individualApp.activeUninstallScript -Force
				}
			}
		} catch {
			Write-Verbose "  Failed to fetch individual app for script detection: $_"
		}
	}
	
	Write-Verbose "Win32App '$($AppData.displayName)': hasInstallScript=$hasInstallScript, hasUninstallScript=$hasUninstallScript"
	
	if ($hasInstallScript -or $hasUninstallScript) {
		# New script-based installation
		if ($hasInstallScript) {
			if ($ExtendedReport) {
				# Fetch script content
				try {
					$scriptId = $AppData.activeInstallScript.targetId
					$contentVersion = if ($AppData.committedContentVersion) { $AppData.committedContentVersion } else { '1' }
					$scriptUrl = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($AppData.id)/microsoft.graph.win32LobApp/contentVersions/$contentVersion/scripts/$scriptId`?`$select=id,displayName,content,state,microsoft.graph.win32LobAppInstallPowerShellScript/enforceSignatureCheck,microsoft.graph.win32LobAppInstallPowerShellScript/runAs32Bit"
					$scriptData = Invoke-MGGraphGetRequestWithMSGraphAllPages $scriptUrl
					
					if ($scriptData) {
						$output += "Install script : $($scriptData.displayName)"
						if ($scriptData.content) {
							$decodedScript = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($scriptData.content))
							$output += "  Enforce signature check : $(if ($scriptData.enforceSignatureCheck) { 'Yes' } else { 'No' })"
							$output += "  Run as 32-bit : $(if ($scriptData.runAs32Bit) { 'Yes' } else { 'No' })"
							$output += "`n  Install Script Content :"
							$output += "  $decodedScript"
						}
					}
				} catch {
					$output += "Install script : Configured (ID: $($AppData.activeInstallScript.targetId))"
					Write-Verbose "Failed to fetch install script content: $_"
				}
			} else {
				$output += "Install script : Configured (use ExtendedReport to view content)"
				Write-Verbose "  Added to output: Install script : Configured (use ExtendedReport to view content)"
			}
		} else {
			# No install script, show traditional command if present
			if ($AppData.installCommandLine) {
				$output += "Install command : $($AppData.installCommandLine)"
				Write-Verbose "  Added to output: Install command : $($AppData.installCommandLine)"
			}
		}
		
		if ($hasUninstallScript) {
			if ($ExtendedReport) {
				# Fetch script content
				try {
					$scriptId = $AppData.activeUninstallScript.targetId
					$contentVersion = if ($AppData.committedContentVersion) { $AppData.committedContentVersion } else { '1' }
					$scriptUrl = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($AppData.id)/microsoft.graph.win32LobApp/contentVersions/$contentVersion/scripts/$scriptId`?`$select=id,displayName,content,state,microsoft.graph.win32LobAppUninstallPowerShellScript/enforceSignatureCheck,microsoft.graph.win32LobAppUninstallPowerShellScript/runAs32Bit"
					$scriptData = Invoke-MGGraphGetRequestWithMSGraphAllPages $scriptUrl
					
					if ($scriptData) {
						$output += "Uninstall script : $($scriptData.displayName)"
						if ($scriptData.content) {
							$decodedScript = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($scriptData.content))
							$output += "  Enforce signature check : $(if ($scriptData.enforceSignatureCheck) { 'Yes' } else { 'No' })"
							$output += "  Run as 32-bit : $(if ($scriptData.runAs32Bit) { 'Yes' } else { 'No' })"
							$output += "`n  Uninstall Script Content :"
							$output += "  $decodedScript"
						}
					}
				} catch {
					$output += "Uninstall script : Configured (ID: $($AppData.activeUninstallScript.targetId))"
					Write-Verbose "Failed to fetch uninstall script content: $_"
				}
			} else {
				$output += "Uninstall script : Configured (use ExtendedReport to view content)"
			}
		} else {
			# No uninstall script, show traditional command if present
			if ($AppData.uninstallCommandLine) {
				$output += "Uninstall command : $($AppData.uninstallCommandLine)"
			}
		}
	} else {
		# No scripts - use traditional command-line installation
		if ($AppData.installCommandLine) {
			$output += "Install command : $($AppData.installCommandLine)"
		}
		if ($AppData.uninstallCommandLine) {
			$output += "Uninstall command : $($AppData.uninstallCommandLine)"
		}
	}
	if ($AppData.installExperience) {
		$runAs = switch ($AppData.installExperience.runAsAccount) {
			'system' { 'System' }
			'user' { 'User' }
			Default { $AppData.installExperience.runAsAccount }
		}
		$output += "Install behavior : $runAs"
		
		$restartBehavior = switch ($AppData.installExperience.deviceRestartBehavior) {
			'suppress' { 'No specific action' }
			'allow' { 'App install may force a device restart' }
			'basedOnReturnCode' { 'Determine behavior based on return codes' }
			'force' { 'Intune will force mandatory device restart' }
			Default { $AppData.installExperience.deviceRestartBehavior }
		}
		$output += "Device restart behavior : $restartBehavior"
		
		if ($AppData.installExperience.maxRunTimeInMinutes) {
			$output += "Installation time required (mins) : $($AppData.installExperience.maxRunTimeInMinutes)"
		}
	}
	
	# Return codes
	if ($AppData.returnCodes -and $AppData.returnCodes.Count -gt 0) {
		$output += "`nReturn codes :"
		foreach ($returnCode in $AppData.returnCodes) {
			$typeDisplay = switch ($returnCode.type) {
				'success' { 'Success' }
				'softReboot' { 'Soft reboot' }
				'hardReboot' { 'Hard reboot' }
				'retry' { 'Retry' }
				'failed' { 'Failed' }
				Default { $returnCode.type }
			}
			$output += "  $($returnCode.returnCode) - $typeDisplay"
		}
	}
	
	# Requirements
	$output += "`n--- Requirements ---"
	
	# Operating system architecture
	if ($AppData.applicableArchitectures) {
		$archDisplay = $AppData.applicableArchitectures -replace 'x64', 'x64' -replace 'x86', 'x86' -replace 'arm64', 'arm64'
		$output += "Check operating system architecture : $archDisplay"
	}
	
	# Minimum OS
	if ($AppData.minimumSupportedWindowsRelease) {
		$osVersion = switch ($AppData.minimumSupportedWindowsRelease) {
			'1607' { 'Windows 10 1607' }
			'1703' { 'Windows 10 1703' }
			'1709' { 'Windows 10 1709' }
			'1803' { 'Windows 10 1803' }
			'1809' { 'Windows 10 1809' }
			'1903' { 'Windows 10 1903' }
			'1909' { 'Windows 10 1909' }
			'2004' { 'Windows 10 2004' }
			'2H20' { 'Windows 10 20H2' }
			'21H1' { 'Windows 10 21H1' }
			Default { "Windows 10 $($AppData.minimumSupportedWindowsRelease)" }
		}
		$output += "Minimum operating system : $osVersion"
	}
	
	# Additional requirements
	$hasAdditionalReqs = $false
	if ($AppData.minimumFreeDiskSpaceInMB) {
		$output += "Disk space required (MB) : $($AppData.minimumFreeDiskSpaceInMB)"
		$hasAdditionalReqs = $true
	}
	if ($AppData.minimumMemoryInMB) {
		$output += "Physical memory required (MB) : $($AppData.minimumMemoryInMB)"
		$hasAdditionalReqs = $true
	}
	if ($AppData.minimumNumberOfProcessors) {
		$output += "Minimum number of logical processors required : $($AppData.minimumNumberOfProcessors)"
		$hasAdditionalReqs = $true
	}
	if ($AppData.minimumCpuSpeedInMHz) {
		$output += "Minimum CPU speed required (MHz) : $($AppData.minimumCpuSpeedInMHz)"
		$hasAdditionalReqs = $true
	}
	
	# Custom requirement rules
	# Win32LobApp uses requirementRules, Win32CatalogApp uses rules array with ruleType='requirement'
	$requirementRulesToProcess = @()
	if ($AppData.requirementRules -and $AppData.requirementRules.Count -gt 0) {
		$requirementRulesToProcess = $AppData.requirementRules
	} elseif ($AppData.rules -and $AppData.rules.Count -gt 0) {
		$requirementRulesToProcess = $AppData.rules | Where-Object { $_.ruleType -eq 'requirement' }
	}
	
	if ($requirementRulesToProcess -and $requirementRulesToProcess.Count -gt 0) {
		$hasAdditionalReqs = $true
		foreach ($rule in $requirementRulesToProcess) {
			switch ($rule.'@odata.type') {
				'#microsoft.graph.win32LobAppFileSystemRequirement' {
					$output += "`nFile or folder requirement :"
					$output += "  Path : $($rule.path)"
					if ($rule.fileOrFolderName) {
						$output += "  File or folder name : $($rule.fileOrFolderName)"
					}
					$operatorText = switch ($rule.operator) {
						'notConfigured' { 'Not configured' }
						'exists' { 'Exists' }
						'modifiedDate' { 'Modified date' }
						'createdDate' { 'Created date' }
						'version' { 'Version' }
						'sizeInMB' { 'Size in MB' }
						Default { $rule.operator }
					}
					$output += "  Detection type : $operatorText"
					if ($rule.comparisonValue) {
						$output += "  Value : $($rule.comparisonValue)"
					}
				}
				'#microsoft.graph.win32LobAppRegistryRequirement' {
					$output += "`nRegistry requirement :"
					$keyPath = switch ($rule.keyPath) {
						{ $_ -match '^HKEY_LOCAL_MACHINE' } { $_ -replace 'HKEY_LOCAL_MACHINE', 'HKLM' }
						{ $_ -match '^HKEY_CURRENT_USER' } { $_ -replace 'HKEY_CURRENT_USER', 'HKCU' }
						Default { $rule.keyPath }
					}
					$output += "  Key path : $keyPath"
					if ($rule.valueName) {
						$output += "  Value name : $($rule.valueName)"
					}
					$operatorText = switch ($rule.operator) {
						'notConfigured' { 'Not configured' }
						'exists' { 'Key exists' }
						'doesNotExist' { 'Key does not exist' }
						'string' { 'String comparison' }
						'integer' { 'Integer comparison' }
						'version' { 'Version comparison' }
						Default { $rule.operator }
					}
					$output += "  Detection type : $operatorText"
					if ($rule.comparisonValue) {
						$output += "  Value : $($rule.comparisonValue)"
					}
				}
				'#microsoft.graph.win32LobAppPowerShellScriptRequirement' {
					$output += "`nPowerShell script requirement :"
					$output += "  Display name : $($rule.displayName)"
					$output += "  Enforce signature check : $(if ($rule.enforceSignatureCheck) { 'Yes' } else { 'No' })"
					$output += "  Run as 32-bit on 64-bit : $(if ($rule.runAs32Bit) { 'Yes' } else { 'No' })"
					if ($rule.scriptContent) {
						try {
							# Decode Base64 script content
							$decodedScript = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($rule.scriptContent))
							$output += "`n  Script content :"
							$output += "  $decodedScript"
						} catch {
							# If decoding fails, show as-is
							$output += "`n  Script content (Base64) :"
							$output += "  $($rule.scriptContent)"
						}
					}
				}
			}
		}
	}
	
	if (-not $hasAdditionalReqs) {
		$output += "No Additional requirement rules"
	}
	
	# Detection rules
	$output += "`n--- Detection rules ---"
	
	# Win32LobApp uses detectionRules, Win32CatalogApp uses rules array with ruleType='detection'
	$detectionRulesToProcess = @()
	if ($AppData.detectionRules -and $AppData.detectionRules.Count -gt 0) {
		$detectionRulesToProcess = $AppData.detectionRules
	} elseif ($AppData.rules -and $AppData.rules.Count -gt 0) {
		$detectionRulesToProcess = $AppData.rules | Where-Object { $_.ruleType -eq 'detection' }
	}
	
	if ($detectionRulesToProcess -and $detectionRulesToProcess.Count -gt 0) {
		$ruleCount = 0
		foreach ($rule in $detectionRulesToProcess) {
			$ruleCount++
			if ($detectionRulesToProcess.Count -gt 1) {
				$output += "`nRule $ruleCount :"
			}
			
			switch ($rule.'@odata.type') {
				{ $_ -in @('#microsoft.graph.win32LobAppFileSystemDetection', '#microsoft.graph.win32LobAppFileSystemRule') } {
					$output += "File or folder detection :"
					$output += "  Path : $($rule.path)"
					if ($rule.fileOrFolderName) {
						$output += "  File or folder name : $($rule.fileOrFolderName)"
					}
					# Use operationType for Win32CatalogApp, detectionType for Win32LobApp
					$typeValue = if ($rule.operationType) { $rule.operationType } else { $rule.detectionType }
					$detectionType = switch ($typeValue) {
						'notConfigured' { 'Not configured' }
						'exists' { 'File or folder exists' }
						'modifiedDate' { 'Modified date' }
						'createdDate' { 'Created date' }
						'version' { 'Version' }
						'sizeInMB' { 'Size (MB)' }
						'sizeInBytes' { 'Size (bytes)' }
						Default { $typeValue }
					}
					$output += "  Detection type : $detectionType"
					# Handle both Win32CatalogApp (operator/comparisonValue) and Win32LobApp (operator/detectionValue)
					$operatorValue = $rule.operator
					$compValue = if ($rule.comparisonValue) { $rule.comparisonValue } else { $rule.detectionValue }
					if ($operatorValue -and $compValue) {
						$operatorText = switch ($operatorValue) {
							'equal' { 'Equal to' }
							'notEqual' { 'Not equal to' }
							'greaterThan' { 'Greater than' }
							'greaterThanOrEqual' { 'Greater than or equal to' }
							'lessThan' { 'Less than' }
							'lessThanOrEqual' { 'Less than or equal to' }
							Default { $operatorValue }
						}
						$output += "  Operator : $operatorText"
						$output += "  Value : $compValue"
					}
					if ($rule.check32BitOn64System) {
						$output += "  Associated with a 32-bit app on 64-bit : Yes"
					}
				}
				{ $_ -in @('#microsoft.graph.win32LobAppRegistryDetection', '#microsoft.graph.win32LobAppRegistryRule') } {
					$output += "Registry detection :"
					$keyPath = switch ($rule.keyPath) {
						{ $_ -match '^HKEY_LOCAL_MACHINE' } { $_ -replace 'HKEY_LOCAL_MACHINE', 'HKLM' }
						{ $_ -match '^HKEY_CURRENT_USER' } { $_ -replace 'HKEY_CURRENT_USER', 'HKCU' }
						Default { $rule.keyPath }
					}
					$output += "  Key path : $keyPath"
					if ($rule.valueName) {
						$output += "  Value name : $($rule.valueName)"
					}
					# Use operationType for Win32CatalogApp, detectionType for Win32LobApp
					$typeValue = if ($rule.operationType) { $rule.operationType } else { $rule.detectionType }
					$detectionType = switch ($typeValue) {
						'notConfigured' { 'Not configured' }
						'exists' { 'Key or value exists' }
						'doesNotExist' { 'Key or value does not exist' }
						'string' { 'String comparison' }
						'integer' { 'Integer comparison' }
						'version' { 'Version comparison' }
						Default { $typeValue }
					}
					$output += "  Detection type : $detectionType"
					# Handle both Win32CatalogApp (operator/comparisonValue) and Win32LobApp (operator/detectionValue)
					$operatorValue = $rule.operator
					$compValue = if ($rule.comparisonValue) { $rule.comparisonValue } else { $rule.detectionValue }
					if ($operatorValue -and $compValue) {
						$operatorText = switch ($operatorValue) {
							'equal' { 'Equal to' }
							'notEqual' { 'Not equal to' }
							'greaterThan' { 'Greater than' }
							'greaterThanOrEqual' { 'Greater than or equal to' }
							'lessThan' { 'Less than' }
							'lessThanOrEqual' { 'Less than or equal to' }
							Default { $operatorValue }
						}
						$output += "  Operator : $operatorText"
						$output += "  Value : $compValue"
					}
					if ($rule.check32BitOn64System) {
						$output += "  Associated with a 32-bit app on 64-bit : Yes"
					}
				}
				'#microsoft.graph.win32LobAppProductCodeDetection' {
					$output += "MSI product code detection :"
					$output += "  Product code : $($rule.productCode)"
					if ($rule.productVersion) {
						$versionOperator = switch ($rule.productVersionOperator) {
							'notConfigured' { 'Not configured' }
							'equal' { 'Equal' }
							'notEqual' { 'Not equal' }
							'greaterThan' { 'Greater than' }
							'greaterThanOrEqual' { 'Greater than or equal' }
							'lessThan' { 'Less than' }
							'lessThanOrEqual' { 'Less than or equal' }
							Default { $rule.productVersionOperator }
						}
						$output += "  Product version operator : $versionOperator"
						$output += "  Product version : $($rule.productVersion)"
					}
				}
				'#microsoft.graph.win32LobAppPowerShellScriptDetection' {
					$output += "PowerShell script detection :"
					$output += "  Enforce signature check : $(if ($rule.enforceSignatureCheck) { 'Yes' } else { 'No' })"
					$output += "  Run as 32-bit on 64-bit : $(if ($rule.runAs32Bit) { 'Yes' } else { 'No' })"
					if ($rule.scriptContent) {
						try {
							# Decode Base64 script content
							$decodedScript = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($rule.scriptContent))
							$output += "`n  Script content :"
							$output += "  $decodedScript"
						} catch {
							# If decoding fails, show as-is
							$output += "`n  Script content (Base64) :"
							$output += "  $($rule.scriptContent)"
						}
					}
				}
			}
		}
	}
	
	# MSI Information (if available)
	if ($AppData.msiInformation) {
		$output += "`n--- MSI Information ---"
		if ($AppData.msiInformation.productName) {
			$output += "Product name : $($AppData.msiInformation.productName)"
		}
		if ($AppData.msiInformation.publisher) {
			$output += "Publisher : $($AppData.msiInformation.publisher)"
		}
		if ($AppData.msiInformation.productVersion) {
			$output += "Product version : $($AppData.msiInformation.productVersion)"
		}
		if ($AppData.msiInformation.productCode) {
			$output += "Product code : $($AppData.msiInformation.productCode)"
		}
		if ($AppData.msiInformation.upgradeCode) {
			$output += "Upgrade code : $($AppData.msiInformation.upgradeCode)"
		}
		$packageType = switch ($AppData.msiInformation.packageType) {
			'perMachine' { 'Per-machine' }
			'perUser' { 'Per-user' }
			'dualPurpose' { 'Dual purpose' }
			Default { $AppData.msiInformation.packageType }
		}
		$output += "Package type : $packageType"
		$output += "Requires reboot : $(if ($AppData.msiInformation.requiresReboot) { 'Yes' } else { 'No' })"
	}

	return ($output -join "`n")
}

function ConvertTo-ReadableMacOSDmgApp {
	param($AppData)

	if (-not $AppData) {
		return $null
	}
	
	$odataType = $AppData.'@odata.type'
	if ($odataType -ne '#microsoft.graph.macOSDmgApp') {
		return $null
	}

	$output = @()
	
	# App Information
	if ($AppData.description) {
		$output += "Description : $($AppData.description)"
	}
	if ($AppData.publisher) {
		$output += "Publisher : $($AppData.publisher)"
	}
	if ($AppData.primaryBundleVersion) {
		$output += "Bundle Version : $($AppData.primaryBundleVersion)"
	}
	if ($AppData.primaryBundleId) {
		$output += "Bundle ID : $($AppData.primaryBundleId)"
	}
	
	# File Information
	$output += "`n--- File Information ---"
	if ($AppData.fileName) {
		$output += "File name : $($AppData.fileName)"
	}
	if ($AppData.size) {
		$sizeGB = [math]::Round($AppData.size / 1GB, 2)
		$sizeMB = [math]::Round($AppData.size / 1MB, 2)
		if ($sizeGB -ge 1) {
			$output += "File size : $sizeGB GB"
		} else {
			$output += "File size : $sizeMB MB"
		}
	}
	
	# Detection Settings
	$output += "`n--- Detection ---"
	if ($AppData.ignoreVersionDetection -ne $null) {
		$output += "Ignore app version : $(if ($AppData.ignoreVersionDetection) { 'Yes' } else { 'No' })"
	}
	
	if ($AppData.includedApps -and $AppData.includedApps.Count -gt 0) {
		$output += "`nIncluded Apps:"
		foreach ($includedApp in $AppData.includedApps) {
			$output += "  Bundle ID : $($includedApp.bundleId)"
			if ($includedApp.bundleVersion) {
				$output += "  Version   : $($includedApp.bundleVersion)"
			}
		}
	}
	
	# Minimum OS Requirements
	if ($AppData.minimumSupportedOperatingSystem) {
		$output += "`n--- Minimum macOS Version ---"
		$minOS = $AppData.minimumSupportedOperatingSystem
		$osVersions = @(
			@{Name='macOS 10.7 (Lion)'; Key='v10_7'},
			@{Name='macOS 10.8 (Mountain Lion)'; Key='v10_8'},
			@{Name='macOS 10.9 (Mavericks)'; Key='v10_9'},
			@{Name='macOS 10.10 (Yosemite)'; Key='v10_10'},
			@{Name='macOS 10.11 (El Capitan)'; Key='v10_11'},
			@{Name='macOS 10.12 (Sierra)'; Key='v10_12'},
			@{Name='macOS 10.13 (High Sierra)'; Key='v10_13'},
			@{Name='macOS 10.14 (Mojave)'; Key='v10_14'},
			@{Name='macOS 10.15 (Catalina)'; Key='v10_15'},
			@{Name='macOS 11.0 (Big Sur)'; Key='v11_0'},
			@{Name='macOS 12.0 (Monterey)'; Key='v12_0'},
			@{Name='macOS 13.0 (Ventura)'; Key='v13_0'},
			@{Name='macOS 14.0 (Sonoma)'; Key='v14_0'},
			@{Name='macOS 15.0 (Sequoia)'; Key='v15_0'}
		)
		
		$requiredOS = $null
		foreach ($ver in $osVersions) {
			if ($minOS.($ver.Key) -eq $true) {
				$requiredOS = $ver.Name
				break
			}
		}
		
		if ($requiredOS) {
			$output += "Minimum OS : $requiredOS"
		}
	}

	return ($output -join "`n")
}

function ConvertTo-ReadableMacOSPkgApp {
	param($AppData)

	if (-not $AppData) {
		return $null
	}
	
	$odataType = $AppData.'@odata.type'
	if ($odataType -ne '#microsoft.graph.macOSPkgApp') {
		return $null
	}

	$output = @()
	
	# App Information
	if ($AppData.description) {
		$output += "Description : $($AppData.description)"
	}
	if ($AppData.publisher) {
		$output += "Publisher : $($AppData.publisher)"
	}
	if ($AppData.primaryBundleVersion) {
		$output += "Bundle Version : $($AppData.primaryBundleVersion)"
	}
	if ($AppData.primaryBundleId) {
		$output += "Bundle ID : $($AppData.primaryBundleId)"
	}
	
	# File Information
	$output += "`n--- File Information ---"
	if ($AppData.fileName) {
		$output += "File name : $($AppData.fileName)"
	}
	if ($AppData.size) {
		$sizeGB = [math]::Round($AppData.size / 1GB, 2)
		$sizeMB = [math]::Round($AppData.size / 1MB, 2)
		if ($sizeGB -ge 1) {
			$output += "File size : $sizeGB GB"
		} else {
			$output += "File size : $sizeMB MB"
		}
	}
	
	# Scripts
	if ($AppData.preInstallScript -or $AppData.postInstallScript) {
		$output += "`n--- Install Scripts ---"
		if ($AppData.preInstallScript) {
			$output += "Pre-install script : Configured"
			if ($AppData.preInstallScript.scriptContent) {
				$output += "  Script content available"
			}
		}
		if ($AppData.postInstallScript) {
			$output += "Post-install script : Configured"
			if ($AppData.postInstallScript.scriptContent) {
				$output += "  Script content available"
			}
		}
	}
	
	# Detection Settings
	$output += "`n--- Detection ---"
	if ($AppData.ignoreVersionDetection -ne $null) {
		$output += "Ignore app version : $(if ($AppData.ignoreVersionDetection) { 'Yes' } else { 'No' })"
	}
	
	if ($AppData.includedApps -and $AppData.includedApps.Count -gt 0) {
		$output += "`nIncluded Apps:"
		foreach ($includedApp in $AppData.includedApps) {
			$output += "  Bundle ID : $($includedApp.bundleId)"
			if ($includedApp.bundleVersion) {
				$output += "  Version   : $($includedApp.bundleVersion)"
			}
		}
	}
	
	# Minimum OS Requirements
	if ($AppData.minimumSupportedOperatingSystem) {
		$output += "`n--- Minimum macOS Version ---"
		$minOS = $AppData.minimumSupportedOperatingSystem
		$osVersions = @(
			@{Name='macOS 10.7 (Lion)'; Key='v10_7'},
			@{Name='macOS 10.8 (Mountain Lion)'; Key='v10_8'},
			@{Name='macOS 10.9 (Mavericks)'; Key='v10_9'},
			@{Name='macOS 10.10 (Yosemite)'; Key='v10_10'},
			@{Name='macOS 10.11 (El Capitan)'; Key='v10_11'},
			@{Name='macOS 10.12 (Sierra)'; Key='v10_12'},
			@{Name='macOS 10.13 (High Sierra)'; Key='v10_13'},
			@{Name='macOS 10.14 (Mojave)'; Key='v10_14'},
			@{Name='macOS 10.15 (Catalina)'; Key='v10_15'},
			@{Name='macOS 11.0 (Big Sur)'; Key='v11_0'},
			@{Name='macOS 12.0 (Monterey)'; Key='v12_0'},
			@{Name='macOS 13.0 (Ventura)'; Key='v13_0'},
			@{Name='macOS 14.0 (Sonoma)'; Key='v14_0'},
			@{Name='macOS 15.0 (Sequoia)'; Key='v15_0'}
		)
		
		$requiredOS = $null
		foreach ($ver in $osVersions) {
			if ($minOS.($ver.Key) -eq $true) {
				$requiredOS = $ver.Name
				break
			}
		}
		
		if ($requiredOS) {
			$output += "Minimum OS : $requiredOS"
		}
	}

	return ($output -join "`n")
}

function ConvertTo-ReadableIosVppApp {
	param($AppData)

	if (-not $AppData) {
		return $null
	}
	
	$odataType = $AppData.'@odata.type'
	if ($odataType -ne '#microsoft.graph.iosVppApp') {
		return $null
	}

	$output = @()
	
	# App Information
	if ($AppData.description) {
		$output += "Description : $($AppData.description)"
	}
	if ($AppData.publisher) {
		$output += "Publisher : $($AppData.publisher)"
	}
	if ($AppData.bundleId) {
		$output += "Bundle ID : $($AppData.bundleId)"
	}
	if ($AppData.informationUrl) {
		$output += "App Store URL : $($AppData.informationUrl)"
	}
	
	# VPP Information
	$output += "`n--- VPP (Volume Purchase Program) ---"
	if ($AppData.vppTokenOrganizationName) {
		$output += "Organization : $($AppData.vppTokenOrganizationName)"
	}
	if ($AppData.vppTokenDisplayName) {
		$output += "VPP Token : $($AppData.vppTokenDisplayName)"
	}
	if ($AppData.vppTokenAppleId) {
		$output += "Apple ID : $($AppData.vppTokenAppleId)"
	}
	if ($AppData.vppTokenAccountType) {
		$accountType = switch ($AppData.vppTokenAccountType) {
			'business' { 'Business' }
			'education' { 'Education' }
			Default { $AppData.vppTokenAccountType }
		}
		$output += "Account Type : $accountType"
	}
	
	# License Information
	$output += "`n--- License Information ---"
	if ($AppData.totalLicenseCount -ne $null) {
		$output += "Total Licenses : $($AppData.totalLicenseCount)"
	}
	if ($AppData.usedLicenseCount -ne $null) {
		$available = $AppData.totalLicenseCount - $AppData.usedLicenseCount
		$output += "Used Licenses : $($AppData.usedLicenseCount)"
		$output += "Available Licenses : $available"
	}
	
	# Licensing Type
	if ($AppData.licensingType) {
		$output += "`nLicensing Support:"
		if ($AppData.licensingType.supportsUserLicensing -or $AppData.licensingType.supportUserLicensing) {
			$output += "  User Licensing : Yes"
		}
		if ($AppData.licensingType.supportsDeviceLicensing -or $AppData.licensingType.supportDeviceLicensing) {
			$output += "  Device Licensing : Yes"
		}
	}
	
	# Applicable Device Type
	if ($AppData.applicableDeviceType) {
		$output += "`n--- Applicable Devices ---"
		$deviceTypes = @()
		if ($AppData.applicableDeviceType.iPad) {
			$deviceTypes += "iPad"
		}
		if ($AppData.applicableDeviceType.iPhoneAndIPod) {
			$deviceTypes += "iPhone and iPod"
		}
		if ($deviceTypes.Count -gt 0) {
			$output += "Device Types : $($deviceTypes -join ', ')"
		}
	}

	return ($output -join "`n")
}

function ConvertTo-ReadableMacOsVppApp {
	param($AppData)

	if (-not $AppData) {
		return $null
	}
	
	$odataType = $AppData.'@odata.type'
	if ($odataType -ne '#microsoft.graph.macOsVppApp') {
		return $null
	}

	$output = @()
	
	# App Information
	if ($AppData.description) {
		$output += "Description : $($AppData.description)"
	}
	if ($AppData.publisher) {
		$output += "Publisher : $($AppData.publisher)"
	}
	if ($AppData.bundleId) {
		$output += "Bundle ID : $($AppData.bundleId)"
	}
	if ($AppData.informationUrl) {
		$output += "App Store URL : $($AppData.informationUrl)"
	}
	
	# VPP Information
	$output += "`n--- VPP (Volume Purchase Program) ---"
	if ($AppData.vppTokenOrganizationName) {
		$output += "Organization : $($AppData.vppTokenOrganizationName)"
	}
	if ($AppData.vppTokenDisplayName) {
		$output += "VPP Token : $($AppData.vppTokenDisplayName)"
	}
	if ($AppData.vppTokenAppleId) {
		$output += "Apple ID : $($AppData.vppTokenAppleId)"
	}
	if ($AppData.vppTokenAccountType) {
		$accountType = switch ($AppData.vppTokenAccountType) {
			'business' { 'Business' }
			'education' { 'Education' }
			Default { $AppData.vppTokenAccountType }
		}
		$output += "Account Type : $accountType"
	}
	
	# License Information
	$output += "`n--- License Information ---"
	if ($AppData.totalLicenseCount -ne $null) {
		$output += "Total Licenses : $($AppData.totalLicenseCount)"
	}
	if ($AppData.usedLicenseCount -ne $null) {
		$available = $AppData.totalLicenseCount - $AppData.usedLicenseCount
		$output += "Used Licenses : $($AppData.usedLicenseCount)"
		$output += "Available Licenses : $available"
	}
	
	# Licensing Type
	if ($AppData.licensingType) {
		$output += "`nLicensing Support:"
		if ($AppData.licensingType.supportsUserLicensing -or $AppData.licensingType.supportUserLicensing) {
			$output += "  User Licensing : Yes"
		}
		if ($AppData.licensingType.supportsDeviceLicensing -or $AppData.licensingType.supportDeviceLicensing) {
			$output += "  Device Licensing : Yes"
		}
	}

	return ($output -join "`n")
}

function ConvertTo-ReadableWebApp {
	param($AppData)

	if (-not $AppData) {
		return $null
	}
	
	$odataType = $AppData.'@odata.type'
	if ($odataType -ne '#microsoft.graph.webApp') {
		return $null
	}

	$output = @()
	
	# App URL - most important
	if ($AppData.appUrl) {
		$output += "App URL : $($AppData.appUrl)"
	}
	
	# App Information
	if ($AppData.description) {
		$output += "Description : $($AppData.description)"
	}
	if ($AppData.publisher) {
		$output += "Publisher : $($AppData.publisher)"
	}
	
	# Browser settings
	if ($AppData.useManagedBrowser -ne $null) {
		$browserSetting = if ($AppData.useManagedBrowser) { 'Yes (Managed Browser required)' } else { 'No (Any browser)' }
		$output += "Use Managed Browser : $browserSetting"
	}
	
	# Additional URLs
	if ($AppData.informationUrl) {
		$output += "Information URL : $($AppData.informationUrl)"
	}
	if ($AppData.privacyInformationUrl) {
		$output += "Privacy URL : $($AppData.privacyInformationUrl)"
	}

	return ($output -join "`n")
}

function ConvertTo-ReadableWinGetApp {
	param($AppData)

	if (-not $AppData) {
		return $null
	}
	
	$odataType = $AppData.'@odata.type'
	if ($odataType -ne '#microsoft.graph.winGetApp') {
		return $null
	}

	$output = @()
	
	# App Information
	if ($AppData.description) {
		$output += "Description : $($AppData.description)"
	}
	if ($AppData.publisher) {
		$output += "`nPublisher : $($AppData.publisher)"
	}
	
	# Package Identifier - most important for WinGet
	if ($AppData.packageIdentifier) {
		$output += "Package Identifier : $($AppData.packageIdentifier)"
	}
	
	# Install Experience
	if ($AppData.installExperience) {
		$output += "`n--- Install Experience ---"
		
		if ($AppData.installExperience.runAsAccount) {
			$runAs = switch ($AppData.installExperience.runAsAccount) {
				'system' { 'System' }
				'user' { 'User' }
				Default { $AppData.installExperience.runAsAccount }
			}
			$output += "Run as account : $runAs"
		}
	}
	
	# Publishing Information
	if ($AppData.publishingState) {
		$output += "`nPublishing State : $($AppData.publishingState)"
	}
	
	# Additional URLs
	if ($AppData.informationUrl) {
		$output += "Information URL : $($AppData.informationUrl)"
	}
	if ($AppData.privacyInformationUrl) {
		$output += "Privacy URL : $($AppData.privacyInformationUrl)"
	}
	
	# App Dependencies
	if ($AppData.dependentAppCount -and $AppData.dependentAppCount -gt 0) {
		$output += "`nDependent Apps : $($AppData.dependentAppCount)"
	}
	if ($AppData.supersedingAppCount -and $AppData.supersedingAppCount -gt 0) {
		$output += "Superseding Apps : $($AppData.supersedingAppCount)"
	}
	if ($AppData.supersededAppCount -and $AppData.supersededAppCount -gt 0) {
		$output += "Superseded Apps : $($AppData.supersededAppCount)"
	}

	return ($output -join "`n")
}

function ConvertTo-ReadableMacOSCustomConfiguration {
	param($PolicyData)

	if (-not $PolicyData) {
		return $null
	}
	
	$odataType = $PolicyData.'@odata.type'
	if ($odataType -ne '#microsoft.graph.macOSCustomConfiguration') {
		return $null
	}

	$output = @()
	
	# Basic Information
	if ($PolicyData.description) {
		$output += "Description : $($PolicyData.description)"
	}
	if ($PolicyData.payloadName) {
		$output += "Payload Name : $($PolicyData.payloadName)"
	}
	if ($PolicyData.payloadFileName) {
		$output += "Filename : $($PolicyData.payloadFileName)"
	}
	if ($PolicyData.deploymentChannel) {
		$channel = switch ($PolicyData.deploymentChannel) {
			'deviceChannel' { 'Device Channel' }
			'userChannel' { 'User Channel' }
			Default { $PolicyData.deploymentChannel }
		}
		$output += "Deployment : $channel"
	}
	
	# Decode and display the payload
	if ($PolicyData.payload) {
		$output += "`n--- Configuration Profile (Decoded) ---`n"
		try {
			# Decode base64 payload
			$decodedBytes = [System.Convert]::FromBase64String($PolicyData.payload)
			$decodedText = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
			
			# Add the decoded XML/plist content
			$output += $decodedText
		}
		catch {
			$output += "Error decoding payload: $_"
		}
	}

	return ($output -join "`n")
}

function ConvertTo-ReadableMacOSCustomAppConfiguration {
	param($PolicyData)

	if (-not $PolicyData) {
		return $null
	}
	
	$odataType = $PolicyData.'@odata.type'
	if ($odataType -ne '#microsoft.graph.macOSCustomAppConfiguration') {
		return $null
	}

	$output = @()
	
	# Basic Information
	if ($PolicyData.description) {
		$output += "Description : $($PolicyData.description)"
	}
	if ($PolicyData.bundleId) {
		$output += "Bundle ID : $($PolicyData.bundleId)"
	}
	if ($PolicyData.fileName) {
		$output += "Filename : $($PolicyData.fileName)"
	}
	
	# Decode and display the configuration XML (plist fragment)
	if ($PolicyData.configurationXml) {
		$output += "`n--- Plist Configuration (Decoded) ---`n"
		try {
			# Decode base64 configuration XML
			$decodedBytes = [System.Convert]::FromBase64String($PolicyData.configurationXml)
			$decodedText = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
			
			# Add the decoded plist fragment
			$output += $decodedText
		}
		catch {
			$output += "Error decoding configuration XML: $_"
		}
	}

	return ($output -join "`n")
}

function ConvertTo-ReadableAppleEnrollmentProfile {
	param($ProfileData)

	if (-not $ProfileData) {
		return $null
	}
	
	$output = @()
	
	# Determine profile type - try multiple methods
	$profileType = 'Unknown'
	
	# Method 1: Check @odata.type
	if ($ProfileData.'@odata.type' -eq '#microsoft.graph.depIOSEnrollmentProfile') {
		$profileType = 'iOS'
	} elseif ($ProfileData.'@odata.type' -eq '#microsoft.graph.depMacOSEnrollmentProfile') {
		$profileType = 'macOS'
	}
	# Method 2: If still unknown, try to detect based on specific properties
	elseif ($ProfileData.fileVaultDisabled -ne $null -or 
	        $ProfileData.iCloudDiagnosticsDisabled -ne $null -or
	        $ProfileData.iCloudStorageDisabled -ne $null -or
	        $ProfileData.registrationDisabled -ne $null -or
	        $ProfileData.skipPrimarySetupAccountCreation -ne $null) {
		# These properties only exist in macOS profiles
		$profileType = 'macOS'
	}
	# Method 3: Check if it has iOS-specific properties
	elseif ($ProfileData.enableSharedIPad -ne $null -or
	        $ProfileData.iTunesPairingMode -ne $null) {
		$profileType = 'iOS'
	}
	
	# Header - Name, Description, Platform
	if ($ProfileData.displayName) {
		$output += "Name: $($ProfileData.displayName)"
	}
	if ($ProfileData.description) {
		$output += "Description: $($ProfileData.description)"
	}
	$output += "Platform: $profileType"
	$output += ""
	
	# Management Settings
	$output += "=== Management Settings ==="
	$output += ""
	$output += "User Affinity & Authentication Method"
	if ($ProfileData.requiresUserAuthentication) {
		$output += "  User affinity: Enroll with User Affinity"
		$authMethod = if ($ProfileData.enableAuthenticationViaCompanyPortal) {
			"Company Portal"
		} else {
			"Setup Assistant with modern authentication"
		}
		$output += "  Authentication Method: $authMethod"
	} else {
		$output += "  User affinity: Enroll without User Affinity"
	}
	$output += ""
	
	$output += "Management Options"
	if ($null -ne $ProfileData.waitForDeviceConfiguredConfirmation) {
		$awaitConfig = if ($ProfileData.waitForDeviceConfiguredConfirmation) { "Yes" } else { "No" }
		$output += "  Await final configuration: $awaitConfig"
	}
	if ($null -ne $ProfileData.profileRemovalDisabled) {
		$lockedEnroll = if ($ProfileData.profileRemovalDisabled) { "Yes" } else { "No" }
		$output += "  Locked enrollment: $lockedEnroll"
	}
	$output += ""
	
	# Setup Assistant
	$output += "=== Setup Assistant ==="
	$output += ""
	
	if ($ProfileData.supportDepartment -or $ProfileData.supportPhoneNumber) {
		$output += "Department"
		if ($ProfileData.supportDepartment) {
			$output += "  $($ProfileData.supportDepartment)"
		}
		$output += "Department Phone"
		if ($ProfileData.supportPhoneNumber) {
			$output += "  $($ProfileData.supportPhoneNumber)"
		}
		$output += ""
	}
	
	$output += "Setup Assistant Screens"
	
	# Create mapping of API keys to UI labels
	$screenMapping = @{
		'Location' = 'Location Services'
		'Restore' = 'Restore'
		'AppleID' = 'Apple ID'
		'TOS' = 'Terms and conditions'
		'Biometric' = 'Touch ID and Face ID'
		'TouchId' = 'Touch ID and Face ID'
		'Payment' = 'Apple Pay'
		'Siri' = 'Siri'
		'Diagnostics' = 'Diagnostics Data'
		'DisplayTone' = 'Display Tone'
		'Privacy' = 'Privacy'
		'ScreenTime' = 'Screen Time'
		'Zoom' = 'Zoom'
		'Android' = 'Android'
		'HomeButtonSensitivity' = 'Home Button Sensitivity'
		'iMessageAndFaceTime' = 'iMessage and FaceTime'
		'OnBoarding' = 'OnBoarding'
		'WatchMigration' = 'Watch Migration'
		'Passcode' = 'Passcode'
		'Welcome' = 'Welcome'
		'RestoreCompleted' = 'Restore Completed'
		'UpdateCompleted' = 'Update Completed'
		'DeviceToDeviceMigration' = 'Device to Device Migration'
		'SIMSetup' = 'SIM Setup'
		'Appearance' = 'Appearance'
		'FileVault' = 'FileVault'
		'iCloudDiagnostics' = 'iCloud Diagnostics'
		'iCloudStorage' = 'iCloud Storage'
		'Registration' = 'Registration'
		'Accessibility' = 'Accessibility'
		'UnlockWithWatch' = 'Auto unlock with Apple Watch'
		'Lockdown' = 'Lockdown mode'
		'EnableLockdownMode' = 'Lockdown mode'
		'Wallpaper' = 'Wallpaper'
		'SoftwareUpdate' = 'Software Update'
		'TermsOfAddress' = 'Terms of Address'
		'Intelligence' = 'Intelligence'
		'Safety' = 'Safety'
		'ActionButton' = 'Action Button'
	}
	
	# Determine which screens are shown or hidden
	$skippedKeys = if ($ProfileData.enabledSkipKeys) { $ProfileData.enabledSkipKeys } else { @() }
	
	# Define the order and screens for each platform type
	if ($profileType -eq 'macOS') {
		# macOS specific screens in order
		$orderedScreens = @(
			'Location', 'Restore', 'AppleID', 'TOS', 'Biometric', 'Payment', 'Siri', 
			'Diagnostics', 'DisplayTone', 'Privacy', 'ScreenTime', 'iCloudDiagnostics', 
			'iCloudStorage', 'Appearance', 'Registration', 'Accessibility', 'UnlockWithWatch',
			'TermsOfAddress', 'Intelligence', 'EnableLockdownMode', 'Wallpaper', 'FileVault'
		)
	} else {
		# iOS/iPadOS screens in order (based on screenshot)
		$orderedScreens = @(
			'Location', 'Restore', 'AppleID', 'TOS', 'Biometric', 'Payment', 'Siri', 
			'Diagnostics', 'DisplayTone', 'Privacy', 'ScreenTime', 'Zoom', 'Android',
			'HomeButtonSensitivity', 'iMessageAndFaceTime', 'OnBoarding', 'WatchMigration',
			'Passcode', 'Welcome', 'RestoreCompleted', 'UpdateCompleted', 
			'DeviceToDeviceMigration', 'SIMSetup', 'Appearance'
		)
	}
	
	# Output each screen with Show/Hide status
	foreach ($screen in $orderedScreens) {
		if ($screenMapping.ContainsKey($screen)) {
			$screenLabel = $screenMapping[$screen]
			$status = if ($skippedKeys -contains $screen) { "Hide" } else { "Show" }
			$output += "  $screenLabel`: $status"
		}
	}
	
	# Add additional screens that might be present but not in our ordered list
	foreach ($screen in $skippedKeys) {
		if ($screenMapping.ContainsKey($screen) -and $orderedScreens -notcontains $screen) {
			$screenLabel = $screenMapping[$screen]
			$output += "  $screenLabel`: Hide"
		}
	}
	
	# Add OS Showcase and App Store for macOS
	if ($profileType -eq 'macOS') {
		$output += "  OS showcase: Show"
		$output += "  App Store: Show"
	}
	
	$output += ""
	
	# Account Settings (macOS only)
	if ($profileType -eq 'macOS') {
		$output += "=== Account Settings ==="
		$output += ""
		
		$output += "Local administrator account"
		$createAdmin = if ($ProfileData.enableRestrictEditing) { "Yes" } else { "No" }
		$output += "  Create a local admin account: $createAdmin"
		
		if ($ProfileData.enableRestrictEditing) {
			if ($ProfileData.adminAccountUserName) {
				$output += "  Admin account username: $($ProfileData.adminAccountUserName)"
			}
			if ($ProfileData.adminAccountFullName) {
				$output += "  Admin account full name: $($ProfileData.adminAccountFullName)"
			}
			if ($null -ne $ProfileData.hideAdminAccount) {
				$hideAdmin = if ($ProfileData.hideAdminAccount) { "Yes" } else { "No" }
				$output += "  Hide in Users & Groups: $hideAdmin"
			}
			if ($ProfileData.depProfileAdminAccountPasswordRotationSetting) {
				$output += "  Admin account password rotation period (days): $($ProfileData.depProfileAdminAccountPasswordRotationSetting)"
			} else {
				$output += "  Admin account password rotation period (days): No Admin account password rotation period (days)"
			}
		}
		$output += ""
		
		$output += "Local user account"
		$createLocal = if ($ProfileData.skipPrimarySetupAccountCreation -eq $false) { "Yes" } else { "No" }
		$output += "  Create a local primary account: $createLocal"
		
		if ($ProfileData.skipPrimarySetupAccountCreation -eq $false) {
			$accountType = if ($ProfileData.setPrimarySetupAccountAsRegularUser) { "Standard" } else { "Administrator" }
			$output += "  Account type: $accountType"
			
			if ($null -ne $ProfileData.dontAutoPopulatePrimaryAccountInfo) {
				$prefill = if ($ProfileData.dontAutoPopulatePrimaryAccountInfo) { "No" } else { "Yes" }
				$output += "  Prefill account info: $prefill"
			}
			
			if ($ProfileData.primaryAccountFullName) {
				$output += "  Primary account name: $($ProfileData.primaryAccountFullName)"
			}
			if ($ProfileData.primaryAccountUserName) {
				$output += "  Primary account full name: $($ProfileData.primaryAccountUserName)"
			}
			
			if ($null -ne $ProfileData.enableRestrictEditing) {
				$restrictEdit = if ($ProfileData.enableRestrictEditing) { "Yes" } else { "No" }
				$output += "  Restrict editing: $restrictEdit"
			}
		}
	}
	
	return ($output -join "`n")
}

function Update-QuickFilters {
	$maxDevicesLabel = "(Max $GraphAPITop devices)"
	$filters = @()

	$filters += [pscustomobject]@{
		QuickFilterName = 'Search by deviceName, serialNumber, emailAddress, OS or id'
		QuickFilterGraphAPIFilter = $null
	}

	$AddSyncFilter = {
		param([string]$Label,[datetime]$Since)
		$timestamp = $Since.ToUniversalTime().ToString('yyyy-MM-ddTHH\:mm\:ss.000Z')
		return [pscustomobject]@{
			QuickFilterName = ("{0,-55} {1,-20}" -f $Label, $maxDevicesLabel)
			QuickFilterGraphAPIFilter = "(lastSyncDateTime gt $timestamp)&`$top=$GraphAPITop"
		}
	}

	$AddEnrollFilter = {
		param([string]$Label,[datetime]$Since)
		$timestamp = $Since.ToUniversalTime().ToString('yyyy-MM-ddTHH\:mm\:ss.000Z')
		return [pscustomobject]@{
			QuickFilterName = ("{0,-55} {1,-20}" -f $Label, $maxDevicesLabel)
			QuickFilterGraphAPIFilter = "(enrolleddatetime gt $timestamp)&`$top=$GraphAPITop"
		}
	}

	$syncLabels = @(
		@('Quick filter: Devices Synced     in last 15 minutes', (Get-Date).AddMinutes(-15)),
		@('Quick filter: Devices Synced     in last  1 hour', (Get-Date).AddHours(-1)),
		@('Quick filter: Devices Synced     in last 24 hours', (Get-Date).AddHours(-24)),
		@('Quick filter: Devices Synced     today (since midnight)', (Get-Date -Hour 0 -Minute 0 -Second 0)),
		@('Quick filter: Devices Synced     in last  7 days', ((Get-Date).Date.AddDays(-7))),
		@('Quick filter: Devices Synced     in last 30 days', ((Get-Date).Date.AddDays(-30)))
	)
	foreach ($entry in $syncLabels) {
		$filters += & $AddSyncFilter $entry[0] $entry[1]
	}

	$enrollLabels = @(
		@('Quick filter: Devices Enrolled   in last 15 minutes', (Get-Date).AddMinutes(-15)),
		@('Quick filter: Devices Enrolled   in last  1 hour', (Get-Date).AddHours(-1)),
		@('Quick filter: Devices Enrolled   today (since midnight)', (Get-Date -Hour 0 -Minute 0 -Second 0)),
		@('Quick filter: Devices Enrolled   in last  7 days', ((Get-Date).Date.AddDays(-7))),
		@('Quick filter: Devices Enrolled   in last 30 days', ((Get-Date).Date.AddDays(-30)))
	)
	foreach ($entry in $enrollLabels) {
		$filters += & $AddEnrollFilter $entry[0] $entry[1]
	}

	$filters += [pscustomobject]@{
		QuickFilterName = ("{0,-32} {1,-22} {2,-20}" -f 'Quick filter: Compliance','Compliant',$maxDevicesLabel)
		QuickFilterGraphAPIFilter = "(complianceState eq 'compliant')&`$top=$GraphAPITop"
	}
	$filters += [pscustomobject]@{
		QuickFilterName = ("{0,-32} {1,-22} {2,-20}" -f 'Quick filter: Compliance','Non-compliant',$maxDevicesLabel)
		QuickFilterGraphAPIFilter = "(complianceState eq 'noncompliant')&`$top=$GraphAPITop"
	}
	$filters += [pscustomobject]@{
		QuickFilterName = ("{0,-32} {1,-22} {2,-20}" -f 'Quick filter: Compliance','Unknown',$maxDevicesLabel)
		QuickFilterGraphAPIFilter = "(complianceState eq 'unknown')&`$top=$GraphAPITop"
	}
	$filters += [pscustomobject]@{
		QuickFilterName = ("{0,-32} {1,-22} {2,-20}" -f 'Quick filter: Ownership','Company devices',$maxDevicesLabel)
		QuickFilterGraphAPIFilter = "(ownerType eq 'company')&`$top=$GraphAPITop"
	}
	$filters += [pscustomobject]@{
		QuickFilterName = ("{0,-32} {1,-22} {2,-20}" -f 'Quick filter: Ownership','Personal devices',$maxDevicesLabel)
		QuickFilterGraphAPIFilter = "(ownerType eq 'personal')&`$top=$GraphAPITop"
	}

	return $filters
}

function Search-ManagedDevices {
	param(
		[string]$SearchString = '',
		[PSCustomObject]$QuickFilter,
		[switch]$IsQuickFilter
	)
	$results = @()
	if ($IsQuickFilter -and $QuickFilter) {
		$filter = $QuickFilter.QuickFilterGraphAPIFilter
		if (-not $filter) { return @() }
		$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=$filter&`$select=id,deviceName,usersLoggedOn,lastSyncDateTime,operatingSystem,deviceType,enrolledDateTime,Manufacturer,Model,SerialNumber,userPrincipalName"
		$url = Fix-UrlSpecialCharacters $url
		$results = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
		$results = $results | Sort-Object -Property deviceName
	} else {
		if (Validate-GUID $SearchString) {
			$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$SearchString"
			$device = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
			if ($device) { $results += $device }
		} else {
			# Search by deviceName
			$query = Fix-UrlSpecialCharacters $SearchString
			$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=contains(deviceName,%27$query%27)&`$select=id,deviceName,usersLoggedOn,lastSyncDateTime,operatingSystem,deviceType,enrolledDateTime,Manufacturer,Model,SerialNumber,userPrincipalName&`$Top=$GraphAPITop"
			$results = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
			$results = $results | Sort-Object -Property deviceName
			if (($SearchString -like '*@*.*') -or ($SearchString -like '*%40*.*')) {
				$userUrl = "https://graph.microsoft.com/beta/users?`$filter=userPrincipalName%20eq%20'$query'&`$select=id,mail,userPrincipalName"
				$azureUser = Invoke-MGGraphGetRequestWithMSGraphAllPages $userUrl
				if ($azureUser -and -not ($azureUser -is [array])) {
					$deviceUrl = "https://graph.microsoft.com/beta/users/$($azureUser.id)/getLoggedOnManagedDevices?`$select=id,deviceName,usersLoggedOn,lastSyncDateTime,operatingSystem,deviceType,enrolledDateTime,Manufacturer,Model,SerialNumber,userPrincipalName"
					$userDevices = Invoke-MGGraphGetRequestWithMSGraphAllPages $deviceUrl
					$results += $userDevices
				}
			}
		}
	}
	$results = [array]$results
	foreach ($device in $results) {
		$lastSyncDays = if ($device.lastSyncDateTime) { (New-TimeSpan $device.lastSyncDateTime).Days } else { 999 }
		$device | Add-Member -NotePropertyName 'searchStringDeviceProperty' -NotePropertyValue ("{0,-25} {1,-10} {2,4} {3,8}" -f $device.deviceName, 'Last sync', $lastSyncDays, 'days ago') -Force
		$toolTip = ($device | Select-Object deviceName,userPrincipalName,operatingSystem,Manufacturer,Model,SerialNumber | Format-List | Out-String).Trim()
		$device | Add-Member -NotePropertyName 'SearchResultToolTip' -NotePropertyValue $toolTip -Force
	}
	return $results
}

function Get-CheckedInUsersInfo {
	param([PSObject]$SelectedUser)
	$usersLoggedOnString = ''
	$latestUser = $null
	$latestGroups = @()
	$collection = @()
	$orderedUsers = $script:IntuneManagedDevice.usersLoggedOn | Sort-Object -Property lastLogOnDateTime -Descending
	foreach ($loggedOn in $orderedUsers) {
		if (-not (Validate-GUID $loggedOn.userId)) { continue }
		if (-not $latestUser) {
			if ($SelectedUser) {
				if ($loggedOn.userId -ne $SelectedUser.id) {
					continue
				}
			}
			if ($script:PrimaryUser -and $loggedOn.userId -eq $script:PrimaryUser.id) {
				$latestUser = $script:PrimaryUser
			} else {
				$userUrl = "https://graph.microsoft.com/beta/users/$($loggedOn.userId)?`$select=*"
				$latestUser = Invoke-MGGraphGetRequestWithMSGraphAllPages $userUrl
			}
			$groupUrl = "https://graph.microsoft.com/beta/users/$($latestUser.id)/memberOf?_=1577625591876"
			$latestGroups = Invoke-MGGraphGetRequestWithMSGraphAllPages $groupUrl
			if ($latestGroups) {
				$latestGroups = Add-AzureADGroupGroupTypeExtraProperties $latestGroups
				$latestGroups = Add-AzureADGroupDevicesAndUserMemberCountExtraProperties $latestGroups
			}
		}
		$userUrl = "https://graph.microsoft.com/beta/users/$($loggedOn.userId)?`$select=id,displayName,mail,userPrincipalName"
		$aadUser = Invoke-MGGraphGetRequestWithMSGraphAllPages $userUrl
		$usersLoggedOnString += "$($aadUser.userPrincipalName)`n"
		$usersLoggedOnString += "$(ConvertTo-LocalDateTimeString $loggedOn.lastLogOnDateTime)`n`n"
		$collection += $aadUser
	}
	return [pscustomobject]@{
		LatestUser   = $latestUser
		LatestGroups = $latestGroups
		RecentText   = $usersLoggedOnString.Trim()
		LoggedOn     = $collection
	}
}

function Get-MobileAppAssignments {
	param(
		[Parameter(Mandatory)][string]$UserId,
		[Parameter(Mandatory)][string]$IntuneDeviceId
	)

	$script:AppsAssignmentsObservableCollection = @()
	$script:UnknownAppAssignments = $false

	if (-not $script:AppsWithAssignments) {
		$script:AppsWithAssignments = Get-ApplicationsWithAssignments -ReloadCacheData:$false
	}
	if (-not $UserId -or -not $IntuneDeviceId) {
		return [pscustomobject]@{
			Items                 = @()
			UnknownAssignments    = $false
		}
	}

	$url = "https://graph.microsoft.com/beta/users('$UserId')/mobileAppIntentAndStates('$IntuneDeviceId')"
	$intentResponse = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
	if (-not $intentResponse.mobileAppList) {
		return [pscustomobject]@{
			Items              = @()
			UnknownAssignments = $false
		}
	}

	$copyOfMobileAppList = $intentResponse.mobileAppList
	foreach ($mobileApp in $intentResponse.mobileAppList) {
		$app = $script:AppsWithAssignments | Where-Object { $_.id -eq $mobileApp.applicationId }
		if (-not $app) { continue }

		foreach ($assignment in $app.assignments) {
			$include = $false
			$context = '_unknown'
			$contextToolTip = ''
			$assignmentGroup = 'unknown'
			$assignmentGroupId = ''
			$assignmentGroupMembers = 'N/A'
			$assignmentGroupTooltip = ''
			$membershipType = ''
			$filterDisplayName = ''
			$filterId = ''
			$filterMode = ''
			$filterTooltip = ''

			if ($assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
				$context = 'User'
				$contextToolTip = 'Built-in All Users group'
				$assignmentGroup = 'All Users'
				$include = $true
			}
			elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
				$context = 'Device'
				$contextToolTip = 'Built-in All Devices group'
				$assignmentGroup = 'All Devices'
				$include = $true
			}
			elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget') {
				$group = $script:deviceGroupMemberships | Where-Object { $_.id -eq $assignment.target.groupId }
				if ($group) {
					$context = 'Device'
					$contextToolTip = $script:IntuneManagedDevice.deviceName
					$assignmentGroup = $group.displayName
					$assignmentGroupId = $group.id
					$assignmentGroupTooltip = $group.membershipRule
					$membershipType = $group.YodamiittiCustomMembershipType
					$assignmentGroupMembers = ''
					if ($group.YodamiittiCustomGroupMembersCountDevices -gt 0) {
						$assignmentGroupMembers += "$($group.YodamiittiCustomGroupMembersCountDevices) devices "
					}
					if ($group.YodamiittiCustomGroupMembersCountUsers -gt 0) {
						$assignmentGroupMembers += "$($group.YodamiittiCustomGroupMembersCountUsers) users "
					}
					$include = $true
				}
				$primaryGroup = $script:PrimaryUserGroupsMemberOf | Where-Object { $_.id -eq $assignment.target.groupId }
				if ($primaryGroup) {
					$context = if ($context -eq 'Device') { '_Device/User' } else { 'User' }
					$contextToolTip = $script:PrimaryUser.userPrincipalName
					$assignmentGroup = $primaryGroup.displayName
					$assignmentGroupId = $primaryGroup.id
					$assignmentGroupTooltip = $primaryGroup.membershipRule
					$membershipType = $primaryGroup.YodamiittiCustomMembershipType
					$assignmentGroupMembers = ''
					if ($primaryGroup.YodamiittiCustomGroupMembersCountDevices -gt 0) {
						$assignmentGroupMembers += "$($primaryGroup.YodamiittiCustomGroupMembersCountDevices) devices "
					}
					if ($primaryGroup.YodamiittiCustomGroupMembersCountUsers -gt 0) {
						$assignmentGroupMembers += "$($primaryGroup.YodamiittiCustomGroupMembersCountUsers) users "
					}
					$include = $true
				}
				$latestGroup = $script:LatestCheckedInUserGroupsMemberOf | Where-Object { $_.id -eq $assignment.target.groupId }
				if ($latestGroup -and $UserId -ne $script:PrimaryUser.id) {
					$context = if ($context -eq 'Device') { '_Device/User' } else { 'User' }
					$contextToolTip = $script:LatestCheckedInUser.userPrincipalName
					$assignmentGroup = $latestGroup.displayName
					$assignmentGroupId = $latestGroup.id
					$assignmentGroupTooltip = $latestGroup.membershipRule
					$membershipType = $latestGroup.YodamiittiCustomMembershipType
					$assignmentGroupMembers = ''
					if ($latestGroup.YodamiittiCustomGroupMembersCountDevices -gt 0) {
						$assignmentGroupMembers += "$($latestGroup.YodamiittiCustomGroupMembersCountDevices) devices "
					}
					if ($latestGroup.YodamiittiCustomGroupMembersCountUsers -gt 0) {
						$assignmentGroupMembers += "$($latestGroup.YodamiittiCustomGroupMembersCountUsers) users "
					}
					$include = $true
				}
			}

			if (-not $include) { continue }

			$filterId = $assignment.target.deviceAndAppManagementAssignmentFilterId
			if ($filterId) {
				$filter = $script:AllIntuneFilters | Where-Object { $_.id -eq $filterId }
				$filterDisplayName = $filter.displayName
				$filterMode = $assignment.target.deviceAndAppManagementAssignmentFilterType
				if ($filterMode -eq 'none') { $filterMode = '' }
				$filterTooltip = $filter.rule
			}

			$assignmentIntent = $assignment.intent
			$includeExclude = switch ($assignment.target.'@odata.type') {
				'#microsoft.graph.groupAssignmentTarget'      { 'Included' }
				'#microsoft.graph.exclusionGroupAssignmentTarget' { 'Excluded' }
				Default { '' }
			}

			if (($assignmentIntent -eq 'available') -and ($mobileApp.installState -eq 'unknown')) {
				$mobileApp.installState = 'Available for install'
			}
			elseif (($assignmentIntent -eq 'required') -and ($mobileApp.installState -eq 'unknown')) {
				$mobileApp.installState = 'Waiting for install status'
			}

			$displayName = if ($app.licenseType -eq 'offline') { "$($app.displayName) (offline)" } else { $app.displayName }
			$odataType = $app.'@odata.type'.Replace('#microsoft.graph.', '')

			$properties = [ordered]@{
				context                        = [string]$context
				contextToolTip                 = [string]$contextToolTip
				odatatype                      = [string]$odataType
				displayName                    = [string]$displayName
				version                        = [string]$mobileApp.displayVersion
				assignmentIntent               = [string]$assignmentIntent
				IncludeExclude                 = [string]$includeExclude
				assignmentGroup                = [string]$assignmentGroup
				YodamiittiCustomGroupMembers   = [string]$assignmentGroupMembers
				assignmentGroupId              = [string]$assignmentGroupId
				installState                   = [string]$mobileApp.installState
				lastModifiedDateTime           = $app.lastModifiedDateTime
				YodamiittiCustomMembershipType = [string]$membershipType
				id                             = $app.id
				filter                         = [string]$filterDisplayName
				filterId                       = [string]$filterId
				filterMode                     = [string]$filterMode
				filterTooltip                  = [string]$filterTooltip
				AssignmentGroupToolTip         = [string]$assignmentGroupTooltip
				displayNameToolTip             = [string]$app.description
			}
			$script:AppsAssignmentsObservableCollection += [pscustomobject]$properties
		}

		if ($script:AppsAssignmentsObservableCollection | Where-Object { $_.id -eq $mobileApp.applicationId }) {
			$copyOfMobileAppList = $copyOfMobileAppList | Where-Object { $_.applicationId -ne $mobileApp.applicationId }
		} else {
			$script:UnknownAppAssignments = $true
			$assignmentIntent = $mobileApp.mobileAppIntent.Replace('Install','')
			$properties = [ordered]@{
				context                        = '_unknown'
				contextToolTip                 = ''
				odatatype                      = ($app.'@odata.type').Replace('#microsoft.graph.','')
				displayName                    = [string]$app.displayName
				version                        = [string]$mobileApp.displayVersion
				assignmentIntent               = [string]$assignmentIntent
				IncludeExclude                 = ''
				assignmentGroup                = 'unknown (possible nested group or removed assignment)'
				YodamiittiCustomGroupMembers   = 'N/A'
				assignmentGroupId              = ''
				installState                   = [string]$mobileApp.installState
				lastModifiedDateTime           = $app.lastModifiedDateTime
				YodamiittiCustomMembershipType = ''
				id                             = $app.id
				filter                         = ''
				filterId                       = ''
				filterMode                     = ''
				filterTooltip                  = ''
				AssignmentGroupToolTip         = ''
				displayNameToolTip             = ''
			}
			$script:AppsAssignmentsObservableCollection += [pscustomobject]$properties
		}
	}

	return [pscustomobject]@{
		Items              = $script:AppsAssignmentsObservableCollection | Sort-Object -Property context, @{ expression = 'assignmentIntent'; Descending = $true }, IncludeExclude, displayName
		UnknownAssignments = $script:UnknownAppAssignments
	}
}

function Ensure-Directory {
	param([Parameter(Mandatory)][string]$Path)
	if (-not (Test-Path -Path $Path)) {
		New-Item -ItemType Directory -Path $Path -Force | Out-Null
	}
	return (Resolve-Path -Path $Path).Path
}

function Initialize-IntuneSession {
	if ($script:TenantId) { return }
	
	# Check for required Microsoft.Graph.Authentication module
	Write-Host ""
	Write-Host "🔍 Checking for required PowerShell modules..." -ForegroundColor Cyan
	$module = Get-Module -Name Microsoft.Graph.Authentication -ListAvailable
	if (-not $module) {
		Write-Host ""
		Write-Host "❌ Microsoft.Graph.Authentication module is not installed" -ForegroundColor Red
		Write-Host ""
		Write-Host "This script requires the Microsoft Graph Authentication -module." -ForegroundColor Yellow
		Write-Host "Please install it using one of the following commands:" -ForegroundColor Yellow
		Write-Host ""
		Write-Host "  For current user only:" -ForegroundColor Cyan
		Write-Host "  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser" -ForegroundColor White
		Write-Host ""
		Write-Host "  For all users (requires admin):" -ForegroundColor Cyan
		Write-Host "  Install-Module Microsoft.Graph.Authentication -Scope AllUsers" -ForegroundColor White
		Write-Host ""
		throw "Microsoft.Graph.Authentication module is required but not installed."
	}
	
	Write-Host ""
	Write-Host "🔗 Connecting to Microsoft Graph..." -ForegroundColor Cyan
	Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
	$scopes = @(
		'DeviceManagementManagedDevices.Read.All',
		'DeviceManagementApps.Read.All',
		'DeviceManagementConfiguration.Read.All',
		'DeviceManagementServiceConfig.Read.All',
		'DeviceManagementScripts.Read.All',
		'User.Read.All',
		'Group.Read.All',
		'GroupMember.Read.All',
		'Directory.Read.All'
	)
	$null = Connect-MgGraph -Scopes $scopes
	$context = Get-MgContext
	if (-not $context -or -not $context.TenantId) {
		throw 'Unable to determine tenant information from Microsoft Graph context.'
	}
	$script:TenantId = $context.TenantId
	$script:ConnectedUser = $context.Account
	
	# Get tenant display name
	try {
		$orgUrl = "https://graph.microsoft.com/v1.0/organization"
		$org = Invoke-MgGraphRequest -Uri $orgUrl -Method Get -OutputType PSObject
		$script:TenantDisplayName = if ($org.value -and $org.value.Count -gt 0) { $org.value[0].displayName } else { $script:TenantId }
	} catch {
		Write-Warning "Could not retrieve tenant display name: $_"
		$script:TenantDisplayName = $script:TenantId
	}
	
	$cachePath = Join-Path -Path $PSScriptRoot -ChildPath "cache\$($script:TenantId)"
	Ensure-Directory -Path $cachePath | Out-Null
	$script:QuickSearchFilters = Update-QuickFilters
	$script:ReportOutputFolder = Ensure-Directory -Path $script:ReportOutputFolder
	Write-Host ""
	Write-Host "✓ Connected to Microsoft Graph" -ForegroundColor Green
	Write-Host "  Tenant: $($script:TenantDisplayName)" -ForegroundColor Cyan
	Write-Host "  Account: $($script:ConnectedUser)" -ForegroundColor Cyan
	Write-Host ""
}

function Write-DeviceSearchTable {
	param([array]$Devices)
	$format = '{0,3} | {1,-30} | {2,-35} | {3,-12} | {4,-10}'
	Write-Host ""
	Write-Host "📱 Found $($Devices.Count) device(s):" -ForegroundColor Yellow
	Write-Host ""
	Write-Host ($format -f '#','Device','User','OS','Last Sync (days)') -ForegroundColor Cyan -BackgroundColor DarkGray -NoNewline
	Write-Host ""
	Write-Host ('-' * 108) -ForegroundColor DarkGray
	for ($i = 0; $i -lt $Devices.Count; $i++) {
		$device = $Devices[$i]
		$days = if ($device.lastSyncDateTime) { (New-TimeSpan $device.lastSyncDateTime).Days } else { 'n/a' }
		$color = if ($days -eq 'n/a' -or $days -gt 7) { 'Red' } elseif ($days -gt 1) { 'Yellow' } else { 'White' }
		Write-Host ($format -f $i,$device.deviceName,$device.userPrincipalName,$device.operatingSystem,$days) -ForegroundColor $color
	}
	Write-Host ""
}

function Invoke-InteractiveDeviceSelection {
	param(
		[string]$InitialSearch
	)
	$search = $InitialSearch
	while ($true) {
		if (-not $search) {
			Write-Host "🔍 Search for device" -ForegroundColor Green
			$input = Read-Host '   Enter device name/email/serial (? for Quick Search Filters, Q to quit)'
			if ($input -match '^[qQ]$') { return $null }
			if ($input -eq '?') {
				Write-Host ""
				Write-Host "⚡ Quick Filters:" -ForegroundColor Yellow
				Write-Host ""
				for ($idx = 0; $idx -lt $script:QuickSearchFilters.Count; $idx++) {
					$color = if ($idx -eq 0) { 'DarkGray' } else { 'White' }
					Write-Host ("  [{0,2}] {1}" -f $idx,$script:QuickSearchFilters[$idx].QuickFilterName) -ForegroundColor $color
				}
				Write-Host ""
				$choice = Read-Host '   Select filter index'
				if ($choice -match '^\d+$' -and [int]$choice -lt $script:QuickSearchFilters.Count) {
					$filter = $script:QuickSearchFilters[[int]$choice]
					$results = Search-ManagedDevices -QuickFilter $filter -IsQuickFilter
					# Force array to handle PS5.1 vs PS7 differences
					$resultsArray = @($results)
					if ($resultsArray -and $resultsArray.Count -gt 0) {
						Write-DeviceSearchTable -Devices $resultsArray
						if ($resultsArray.Count -eq 1) {
							Write-Host "✓ Auto-selected the only device found" -ForegroundColor Green
							Write-Host ""
							return $resultsArray[0]
						}
						$selection = (Read-Host 'Enter result index or press Enter to search again').Trim()
						[int]$selectedIndex = -1
						if ([int]::TryParse($selection, [ref]$selectedIndex) -and $selectedIndex -ge 0 -and $selectedIndex -lt $resultsArray.Count) {
							return $resultsArray[$selectedIndex]
						}
					} else {
						Write-Warning 'No devices found for the selected quick filter.'
					}
				}
				$search = $null
				continue
			}
			$search = $input
		}
		$results = Search-ManagedDevices -SearchString $search
		# Force array to handle PS5.1 vs PS7 differences
		$resultsArray = @($results)
		if (-not $resultsArray -or $resultsArray.Count -eq 0) {
			Write-Warning 'No devices found. Try another search.'
			$search = $null
			continue
		}
		Write-DeviceSearchTable -Devices $resultsArray
		if ($resultsArray.Count -eq 1) {
			Write-Host "✓ Auto-selected the only device found" -ForegroundColor Green
			Write-Host ""
			return $resultsArray[0]
		}
		$selected = (Read-Host 'Enter result index or press Enter to refine search').Trim()
		[int]$selectedIndex = -1
		if ([int]::TryParse($selected, [ref]$selectedIndex) -and $selectedIndex -ge 0 -and $selectedIndex -lt $resultsArray.Count) {
			return $resultsArray[$selectedIndex]
		}
		$search = $null
	}
}

function Resolve-DeviceId {
	param(
		[string]$PipelineId,
		[string]$SearchText
	)
	if ($PipelineId) { return $PipelineId }
	$device = Invoke-InteractiveDeviceSelection -InitialSearch $SearchText

	Write-Verbose "Selected device: $($device | Format-List | Out-String)"

	if ($device) {
		return $device.id
	}
	return $null
}

function Get-ManagedDeviceSnapshot {
	param([Parameter(Mandatory)][string]$IntuneDeviceId)
	$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($IntuneDeviceId)?`$expand=deviceCategory"
	return Invoke-MGGraphGetRequestWithMSGraphAllPages $url
}

function Get-AdditionalDeviceHardware {
	param([Parameter(Mandatory)][string]$IntuneDeviceId)
	$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($IntuneDeviceId)?`$select=id,hardwareinformation,activationLockBypassCode,iccid,udid,roleScopeTagIds,ethernetMacAddress,processorArchitecture"
	return Invoke-MGGraphGetRequestWithMSGraphAllPages $url
}

function Get-PrimaryUserContext {
	param([PSObject]$Device)

	# Check if userPrincipalName exists (primary user is assigned)
	if (-not $Device.userPrincipalName -or [string]::IsNullOrWhiteSpace($Device.userPrincipalName)) { 
		return $null 
	}

	$userUrl = "https://graph.microsoft.com/beta/users?`$filter=userPrincipalName eq '$($Device.userPrincipalName)'&`$select=*"
	$user = Invoke-MGGraphGetRequestWithMSGraphAllPages $userUrl

	if (-not $user) { return $null }

	# If filter returns array, take first result
	if ($user -is [array]) { $user = $user[0] }

	$groupUrl = "https://graph.microsoft.com/beta/users/$($user.id)/memberOf?_=1577625591876"
	$groups = Invoke-MGGraphGetRequestWithMSGraphAllPages $groupUrl
	if ($groups) {
		$groups = Add-AzureADGroupGroupTypeExtraProperties $groups
		$groups = Add-AzureADGroupDevicesAndUserMemberCountExtraProperties $groups
	}
	return [pscustomobject]@{
		User   = $user
		Groups = $groups
	}
}

function Get-LatestLogonContext {
	param([PSObject]$Device)
	$usersInfo = Get-CheckedInUsersInfo
	return $usersInfo
}

function Get-AzureDeviceContext {
	param([PSObject]$Device)
	if (-not $Device.azureADDeviceId) { return $null }
	$url = "https://graph.microsoft.com/beta/devices?`$filter=deviceId%20eq%20`'$($Device.azureADDeviceId)`'"
	$aadDevice = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
	if ($aadDevice) {
		$groupUrl = "https://graph.microsoft.com/beta/devices/$($aadDevice.id)/transitiveMemberOf?_=1577625591876"
		$deviceGroups = Invoke-MGGraphGetRequestWithMSGraphAllPages $groupUrl
		if ($deviceGroups) {
			$deviceGroups = Add-AzureADGroupGroupTypeExtraProperties $deviceGroups
			$deviceGroups = Add-AzureADGroupDevicesAndUserMemberCountExtraProperties $deviceGroups
		}
		return [pscustomobject]@{
			AzureDevice = $aadDevice
			Groups      = $deviceGroups
		}
	}
	return $null
}

function Get-AutopilotContext {
	param([PSObject]$Device)
	if (-not $Device.autopilotEnrolled -or -not $Device.serialNumber) { return $null }

	$url = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,%27$($Device.serialNumber)%27)"
	$autopilotDevice = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
	if (-not $autopilotDevice) { return $null }

	# Get Device Autopilot Details
	$detailUrl = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities/$($autopilotDevice.id)?`$expand=deploymentProfile,intendedDeploymentProfile"
	$detail = Invoke-MGGraphGetRequestWithMSGraphAllPages $detailUrl

	# Get Autopilot configuration policy details with assignment information
	# Use this as example for uri: https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles/04bfb9da-0144-4788-9691-a06290516807?$expand=assignments
	if ($detail.deploymentProfile -and $detail.deploymentProfile.id) {
		$profileUrl = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles/$($detail.deploymentProfile.id)?`$expand=assignments"
		$autopilotProfile = $null
		$autopilotProfile = Invoke-MGGraphGetRequestWithMSGraphAllPages $profileUrl

		# Add new detail property for deployment profile with assignments
		$detail | Add-Member -NotePropertyName 'DeploymentProfileDetail' -NotePropertyValue $autopilotProfile -Force
	}

	return [pscustomobject]@{
		Device = $autopilotDevice
		Detail = $detail
	}
}

function Get-AutopilotDevicePreparationContext {
	param([PSObject]$Device)

	# Check if device has enrollment profile name
	if (-not $Device.enrollmentProfileName) {
		return $null
	}
	
	$searchTerm = [System.Web.HttpUtility]::UrlEncode("`"$($Device.enrollmentProfileName)`"")
	
	# Template IDs for Device Preparation policies
	$templateIds = @(
		'80d33118-b7b4-40d8-b15f-81be745e053f_1',  # Device Preparation
		'a6157a7f-aa00-42d9-ac82-7d2479f545db_1'   # Device Preparation (alternate)
	)
	
	$devicePrepPolicy = $null
	
	# Search for Device Preparation policy using both template IDs
	foreach ($templateId in $templateIds) {
		$url = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?" + 
			   "`$select=id,name,description,platforms,lastModifiedDateTime,technologies,settingCount,roleScopeTagIds,isAssigned,templateReference,priorityMetaData" +
			   "&`$top=100" +
			   "&`$filter=(technologies has 'enrollment') and (platforms eq 'windows10') and (TemplateReference/templateId eq '$templateId') and (Templatereference/templateFamily eq 'enrollmentConfiguration')" +
			   "&`$search=$searchTerm"
		
		$result = Invoke-MGGraphGetRequestWithMSGraphAllPages $url

		if ($result.id) {
			$devicePrepPolicy = $result
			break
		}
	}

	# Now we found that Device Preparation policy exists

	# Next we are actually fetching full details of the policy including assignments
	# URL example for getting the Device Preparation policy: https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('c2905169-29b3-4580-bbfd-5c0a332d480b')?$expand=settings
	# We also need to make second GET to retrive assignments: https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('c2905169-29b3-4580-bbfd-5c0a332d480b')/assignments

	if ($devicePrepPolicy -and $devicePrepPolicy.id) {
		$detailUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$($devicePrepPolicy.id)')?`$expand=settings"
		$devicePrepPolicyDetail = Invoke-MGGraphGetRequestWithMSGraphAllPages $detailUrl

		# Replace the original policy object with the detailed one
		$devicePrepPolicy = $devicePrepPolicyDetail

		# Get assignments and add to the policy object
		$assignmentsUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$($devicePrepPolicy.id)')/assignments"
		$assignments = Invoke-MGGraphGetRequestWithMSGraphAllPages $assignmentsUrl

		# If we have only 1 assignment then we get the object directly, but we need to wrap it into 'value' array to keep consistent with multiple assignments
		if ($assignments -and -not ($assignments -is [array])) {
			$assignments = [pscustomobject]@{
				value = @($assignments)
			}
		}

		$devicePrepPolicy | Add-Member -NotePropertyName 'assignments' -NotePropertyValue $assignments.value -Force
	}

	if(-not $devicePrepPolicy) {
		return $null
	}

	# Add policy to GUID hashtable
	$devicePrepPolicy = Resolve-AssignmentGroupNames -Object $devicePrepPolicy

	return $devicePrepPolicy
}


function Get-EnrollmentStatusPageContext {
	param([PSObject]$Device)

	if (-not $Device.id) { return $null }
	
	$uri = 'https://graph.microsoft.com/beta/deviceManagement/reports/getEnrollmentConfigurationPoliciesByDevice'
	$body = @"
{
	"search":  "",
	"orderBy":  [

				],
	"select":  [
				   "ProfileName",
				   "UserPrincipalName",
				   "PolicyType",
				   "State",
				   "FilterIds",
				   "Priority",
				   "Target",
				   "LastAppliedTime",
				   "PolicyId"
			   ],
	"filter":  "(DeviceId eq \u0027$($Device.id)\u0027)",
	"skip":  0,
	"top":  50
}
"@

	try {
		$result = Invoke-MGGraphPostRequest -Uri $uri -Body $body

		if (-not $result) {
			return $null
		} else {

			$rows = Objectify_JSON_Schema_and_Data_To_PowershellObjects -ReportData $result
			if (-not $rows) { return $null }

			# DEBUG $rows
			#Write-Host "DEBUG: Enrollment Status Page / Enrollment Restriction Rows:" -ForegroundColor Yellow
			#$rows | ConvertTo-Json -Depth 5 | Set-Clipboard
			#Pause

			# PolicyType values appear as ints in the report rows (eg. 27 = ESP, 22 = Device type enrollment restriction)
			# 
			$espRow = $rows | Where-Object {
				$_.PolicyType -eq 27 -or $_.PolicyType_loc -eq 'Enrollment status page'
			} | Select-Object -First 1

			$restrictionRow = $rows | Where-Object {
				$_.PolicyType -eq 22 -or $_.PolicyType_loc -eq 'Device type enrollment restriction'
			} | Select-Object -First 1

			$espDetail = $null
			if ($espRow -and $espRow.PolicyId) {
				$espUrl = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations/$($espRow.PolicyId)_Windows10EnrollmentCompletionPageConfiguration?`$expand=assignments"
				$espDetail = Invoke-MGGraphGetRequestWithMSGraphAllPages $espUrl
			}

			$restrictionDetail = $null
			if ($restrictionRow -and $restrictionRow.PolicyId) {
				$restrictionUrl = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations/$($restrictionRow.PolicyId)_SinglePlatformRestriction?`$expand=assignments"
				$restrictionDetail = Invoke-MGGraphGetRequestWithMSGraphAllPages $restrictionUrl
			}

			# Keep existing report expectation (ESP object), but also include restrictions.
			return [pscustomobject]@{
				Id     = [string]($espRow.PolicyId)
				Name   = [string]($espRow.ProfileName)
				Detail = $espDetail

				EnrollmentRestriction = if ($restrictionRow) {
					[pscustomobject]@{
						Id     = [string]($restrictionRow.PolicyId)
						Name   = [string]($restrictionRow.ProfileName)
						Detail = $restrictionDetail
					}
				} else { $null }

				RawRows = $rows
			}
		}
	} catch {
		Write-Verbose "Failed to get Enrollment Status Page / Enrollment Restriction:`n$_"
	}
	return $null
}

function Resolve-AssignmentGroupNames {
	param(
		[Parameter(Mandatory)][PSObject]$Object
	)
	
	if (-not $Object) { return $Object }
	
	# Resolve group names in assignments - add displayName inside target object
	if ($Object.assignments) {
		foreach ($assignment in $Object.assignments) {
			if ($assignment.target.groupId) {
				$resolvedName = Get-NameFromGUID -Id $assignment.target.groupId -PreferredProperty 'displayName'
				if ($resolvedName) {
					$assignment.target | Add-Member -NotePropertyName 'displayName' -NotePropertyValue $resolvedName -Force
				}
			}
		}
	}
	
	return $Object
}

function Resolve-EspAssignmentGroupNames {
	param([PSObject]$Esp)
	
	if (-not $Esp -or -not $Esp.Detail) { return $Esp }
	
	# Resolve group names in assignments - add displayName inside target object
	if ($Esp.Detail.assignments) {
		foreach ($assignment in $Esp.Detail.assignments) {
			if ($assignment.target.groupId) {
				$resolvedName = Get-NameFromGUID -Id $assignment.target.groupId -PreferredProperty 'displayName'
				if ($resolvedName) {
					$assignment.target | Add-Member -NotePropertyName 'displayName' -NotePropertyValue $resolvedName -Force
				}
			}
		}
	}
	
	# Replace blocking app GUIDs with display names
	if ($Esp.Detail.selectedMobileAppIds) {
		$resolvedAppNames = @()
		foreach ($appId in $Esp.Detail.selectedMobileAppIds) {
			$resolvedName = Get-NameFromGUID -Id $appId -PreferredProperty 'displayName'
			if ($resolvedName) {
				$resolvedAppNames += $resolvedName
			} else {
				$resolvedAppNames += "Unknown app ($appId)"
			}
		}
		$Esp.Detail.selectedMobileAppIds = $resolvedAppNames
	}
	
	return $Esp
}


function Get-ApplicationAssignmentsContext {
	param(
		[string]$UserId,
		[string]$IntuneDeviceId,
		[switch]$Skip,
		[switch]$ReloadCache
	)
	if ($Skip) { return $null }
	$script:AllIntuneFilters = Download-IntuneFilters
	if ($ReloadCache) {
		$script:AppsWithAssignments = Get-ApplicationsWithAssignments -ReloadCacheData:$true
	}
	$appAssignments = Get-MobileAppAssignments -UserId $UserId -IntuneDeviceId $IntuneDeviceId
	return $appAssignments
}

function Get-ConfigurationPolicyReport {
	param([string]$IntuneDeviceId)

	# Initialize Settings Catalog policy IDs tracking array for Extended Report
	$script:SettingsCatalogPolicyIdsToDownload = @()
	
	# Initialize Custom Configuration policies with encrypted OMA settings tracking
	$script:CustomConfigPoliciesWithSecrets = @{}

	Write-Host "Downloading Intune configuration profiles with assignments…" -ForegroundColor Cyan

	# User Powershell splatting to specify function parameters
	# Limited properties
	#GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$expand=assignments&$select=id,description,createdDateTime,lastModifiedDateTime,name,assignments'
	$Params = @{
		GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$expand=assignments&$select=*'
		jsonCacheFileName = 'configurationPolicies.json'
		ReloadCacheData = $ReloadCache
	}
	$Script:IntuneConfigurationProfilesWithAssignments += Download-IntuneConfigurationProfiles2 @Params

	# Limited properties
	#GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?$expand=assignments&$select=id,description,createdDateTime,lastModifiedDateTime,displayname,assignments'
	$Params = @{
		GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?$expand=assignments&$select=*'
		jsonCacheFileName = 'groupPolicyConfigurations.json'
		ReloadCacheData = $ReloadCache
	}
	$Script:IntuneConfigurationProfilesWithAssignments += Download-IntuneConfigurationProfiles2 @Params

	# Limited properties
	#$GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?$expand=assignments&$select=id,description,createdDateTime,lastModifiedDateTime,displayname,assignments'
	$Params = @{
		GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?$expand=assignments&$select=*'
		jsonCacheFileName = 'deviceConfigurations.json'
		ReloadCacheData = $ReloadCache
	}
	$Script:IntuneConfigurationProfilesWithAssignments += Download-IntuneConfigurationProfiles2 @Params

	# Limited properties
	#raphAPIUrl = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations?$expand=assignments&$select=id,description,createdDateTime,lastModifiedDateTime,displayname,assignments'
	$Params = @{
		GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations?$expand=assignments&$select=*'
		jsonCacheFileName = 'mobileAppConfigurations.json'
		ReloadCacheData = $ReloadCache
	}
	$Script:IntuneConfigurationProfilesWithAssignments += Download-IntuneConfigurationProfiles2 @Params


	$Params = @{
		GraphAPIUrl = 'https://graph.microsoft.com/beta/deviceManagement/intents?$select=*'
		jsonCacheFileName = 'intents.json'
		ReloadCacheData = $ReloadCache
	}
	$Script:IntuneConfigurationProfilesWithAssignments += Download-IntuneConfigurationProfiles2 @Params

	# Add configuration profiles GUIDs to global list for later use
	foreach ($profile in $Script:IntuneConfigurationProfilesWithAssignments) {
		Add-GUIDToHashtable -Object $profile
	}

	Write-Host "Found $($Script:IntuneConfigurationProfilesWithAssignments.Count) configuration profiles"
	Write-Host

	$uri = 'https://graph.microsoft.com/beta/deviceManagement/reports/getConfigurationPoliciesReportForDevice'
	$body = @"
{
    "select":  [
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
    "filter":  "((PolicyBaseTypeName eq \u0027Microsoft.Management.Services.Api.DeviceConfiguration\u0027) or (PolicyBaseTypeName eq \u0027DeviceManagementConfigurationPolicy\u0027) or (PolicyBaseTypeName eq \u0027DeviceConfigurationAdmxPolicy\u0027) or (PolicyBaseTypeName eq \u0027Microsoft.Management.Services.Api.DeviceManagementIntent\u0027)) and (IntuneDeviceId eq \u0027$($IntuneDeviceId)\u0027)",
    "skip":  0,
    "top":  50,
    "orderBy":  [
                    "PolicyName"
                ]
}
"@

	# Download (and convert) Device Configuration Policies report
	Write-Host "Get Intune device Configuration Assignment information"
	$ConfigurationPoliciesReportForDevice = Download-IntunePostTypeReport -Uri $uri -GraphAPIPostBody $body
	Write-Host "Found $($ConfigurationPoliciesReportForDevice.Count) Configuration Assignments"

	$script:ConfigurationsAssignmentsObservableCollection = @()


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
			4 { $ConfigurationPolicyReportState.PolicyStatus = 'Error' }   	   # Device based result ??? - This is unknown but should be error
			5 { $ConfigurationPolicyReportState.PolicyStatus = 'Error' }   	   # User based result?
			6 { $ConfigurationPolicyReportState.PolicyStatus = 'Conflict' }
			Default { }
		}


		if($ConfigurationPolicyReportState.PolicyBaseTypeName -eq 'Microsoft.Management.Services.Api.DeviceManagementIntent') {
			# Endpoint Security templates information does not include assignments
			# So we get assignment information separately to those templates
			#https://graph.microsoft.com/beta/deviceManagement/intents/932d590f-b340-4a7c-b199-048fb98f09b2/assignments

			$url = "https://graph.microsoft.com/beta/deviceManagement/intents/$($ConfigurationPolicyReportState.PolicyId)/assignments"
			$IntuneDeviceConfigurationPolicyAssignments = Invoke-MgGraphGetRequestWithMSGraphAllPages $url
		} else {
			$IntunePolicyObject = $Script:IntuneConfigurationProfilesWithAssignments | Where-Object id -eq $ConfigurationPolicyReportState.PolicyId
			
			$IntuneDeviceConfigurationPolicyAssignments = $IntunePolicyObject.assignments
			$displayNameToolTip = $IntunePolicyObject.description
			
			# Use the actual @odata.type from the policy object if available, instead of the localized UnifiedPolicyType_loc
			if ($IntunePolicyObject.'@odata.type') {
				$odatatype = $IntunePolicyObject.'@odata.type'
			}
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
		# Value can be empty string also so we need to test that also
		if ($odatatype -and -not [string]::IsNullOrWhiteSpace($odatatype)) {
			$odatatype = $odatatype.Replace('#microsoft.graph.', '')
		}
		
		# Map odata type to friendly display names
		$odatatypeDisplayName = switch ($odatatype) {
			'macOSCustomAppConfiguration' { 'Preference file' }
			'macOSCustomConfiguration' { 'Custom' }
			default { $odatatype }
		}
		$odatatype = $odatatypeDisplayName
		
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

				# Group based assignment. We need to get Entra ID Group Name
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
			
				# Track Settings Catalog policy IDs for extended report download
				if ($ExtendedReport -and $ConfigurationPolicyReportState.PolicyBaseTypeName -eq 'DeviceManagementConfigurationPolicy') {
					if ($script:SettingsCatalogPolicyIdsToDownload -notcontains $ConfigurationPolicyReportState.PolicyId) {
						$script:SettingsCatalogPolicyIdsToDownload += $ConfigurationPolicyReportState.PolicyId
						Write-Verbose "Tracking Settings Catalog policy ID for download: $($ConfigurationPolicyReportState.PolicyId) - $($ConfigurationPolicyReportState.PolicyName)"
					}
				}
				
				# Track Custom Configuration policies with encrypted OMA settings for extended report
				if ($ExtendedReport -and $IntunePolicyObject.'@odata.type' -eq '#microsoft.graph.windows10CustomConfiguration' -and $IntunePolicyObject.omaSettings) {
					foreach ($omaSetting in $IntunePolicyObject.omaSettings) {
						if ($omaSetting.isEncrypted -eq $true -and $omaSetting.secretReferenceValueId) {
							if (-not $script:CustomConfigPoliciesWithSecrets.ContainsKey($ConfigurationPolicyReportState.PolicyId)) {
								$script:CustomConfigPoliciesWithSecrets[$ConfigurationPolicyReportState.PolicyId] = @()
							}
							$script:CustomConfigPoliciesWithSecrets[$ConfigurationPolicyReportState.PolicyId] += $omaSetting.secretReferenceValueId
							Write-Verbose "Tracking encrypted OMA setting for policy: $($ConfigurationPolicyReportState.PolicyName) - Secret ID: $($omaSetting.secretReferenceValueId)"
						}
					}
				}

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
			# or assignment is based on nested groups so earlier check did not find Entra ID group where device and/or user is member

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
				# There were assignments in Policy but we could not find which Entra ID group is causing policy to be applied
				Write-Host "Warning: Could not resolve Entra ID Group assignment for Policy $($ConfigurationPolicyReportState.PolicyName)!" -ForegroundColor Yellow
				
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
		if(($script:PrimaryUser) -and ($PolicyInGrid.userPrincipalName -eq $script:PrimaryUser.userPrincipalName)) {
			# Policy UPN value is same than Intune device Primary User and PrimaryUser does exist
			
			# No change needed so continue to next policy in foreach loop
			Continue
		} elseif((-not $script:PrimaryUser) -and ($PolicyInGrid.userPrincipalName -eq $Script:LatestCheckedinUser.UserPrincipalName)) {
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
			if(($script:PrimaryUser) -and ($UserPrincipalNames -contains $script:PrimaryUser.userPrincipalName)) {
				$PolicyInGrid.userPrincipalName = $script:PrimaryUser.userPrincipalName
			} elseif((-not $script:PrimaryUser) -and ($UserPrincipalNames -contains $Script:LatestCheckedinUser.UserPrincipalName)) {
				$PolicyInGrid.userPrincipalName = $Script:LatestCheckedinUser.UserPrincipalName
			} else {
				# If primary user was not listed in duplicate policy entries,
				# use any available UPN from the duplicates (shows which user was logged on when policy was evaluated)
				$nonEmptyUPN = $UserPrincipalNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
				if ($nonEmptyUPN) {
					$PolicyInGrid.userPrincipalName = $nonEmptyUPN
				} else {
					$PolicyInGrid.userPrincipalName = ''
				}
			}
		}
	}

	if($script:ConfigurationsAssignmentsObservableCollectionUnique.Count -gt 1) {
		# ItemsSource works if we are sorting 2 or more objects
		
		return $script:ConfigurationsAssignmentsObservableCollectionUnique | Sort-Object displayName,userPrincipalName
	} else {
		# Only 1 object so we can't do sorting
		# If we try to sort here then our object array breaks and it does not work for ItemsSource
		# Cast as array because otherwise it will fail
		return [array]$script:ConfigurationsAssignmentsObservableCollectionUnique
	}

	#return $ConfigurationPoliciesReportForDevice
}

function Get-RemediationScriptsReport {
	param(
		[string]$IntuneDeviceId,
		[array]$DeviceGroups,
		[array]$PrimaryUserGroups,
		[array]$LatestUserGroups,
		[object]$PrimaryUser,
		[object]$LatestUser
	)

	# Initialize script IDs tracking array for extended report
	$script:ScriptIdsToDownload = @()

	Write-Host "Get Remediation scripts for device..." -ForegroundColor Cyan
	
	$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($IntuneDeviceId)/deviceHealthScriptStates"
	$remediationScriptsForDevice = Invoke-MGGraphGetRequestWithMSGraphAllPages $url
	
	if ($remediationScriptsForDevice) {
		Write-Host "Found $($remediationScriptsForDevice.Count) remediation script states"
	} else {
		Write-Host "No remediation script states found for device"
		$remediationScriptsForDevice = @()
	}
	
	# Download all Remediation scripts with assignments
	$script:RemediationScriptsWithAssignments = Get-RemediationScriptsWithAssignments -ReloadCacheData:$ReloadCache
	Write-Host "Found $($script:RemediationScriptsWithAssignments.Count) total remediation scripts"

	# Download all Platform scripts with assignments
	$script:PlatformScriptsWithAssignments = Get-PlatformScriptsWithAssignments -ReloadCacheData:$ReloadCache
	Write-Host "Found $($script:PlatformScriptsWithAssignments.Count) total platform scripts"

	$results = @()

	foreach ($scriptState in $remediationScriptsForDevice) {
		# Get the script details
		$scriptInfo = $script:RemediationScriptsWithAssignments | Where-Object { $_.id -eq $scriptState.policyId }
		
		# Detection status
		$detectionStatus = switch ($scriptState.detectionState) {
			'success' { 'Without issues' }
			'fail' { 'With issues' }
			'notApplicable' { 'Not applicable' }
			default { $scriptState.detectionState }
		}

		# Remediation status
		$remediationStatus = switch ($scriptState.remediationState) {
			'success' { 'Issue fixed' }
			'fail' { 'With issues' }
			'skipped' { 'Not run' }
			'unknown' { 'Not run' }
			default { $scriptState.remediationState }
		}

		# Status update time
		$lastUpdate = $scriptState.lastStateUpdateDateTime
		$statusUpdateTime = ''
		$statusUpdateTimeTooltip = ''
		if ($lastUpdate) {
			$timespan = New-TimeSpan (Get-Date $lastUpdate) (Get-Date)
			if ($timespan.Days -gt 0) {
				$statusUpdateTime = "$($timespan.Days) days ago"
			} elseif ($timespan.Hours -gt 0) {
				$statusUpdateTime = "$($timespan.Hours) hours ago"
			} else {
				$statusUpdateTime = "$($timespan.Minutes) mins ago"
			}
			$statusUpdateTimeTooltip = (Get-Date $lastUpdate -Format "yyyy-MM-dd HH:mm:ss.fff")
		}

		# Detection tooltip
		$detectionTooltip = $scriptState.preRemediationDetectionScriptOutput
		if ([string]::IsNullOrWhiteSpace($detectionTooltip)) {
			$detectionTooltip = 'No output'
		}

		# Remediation tooltip  
		$remediationTooltip = $scriptState.postRemediationDetectionScriptOutput
		if ([string]::IsNullOrWhiteSpace($remediationTooltip)) {
			$remediationTooltip = 'No output'
		}

		# User principal name (from script state)
		$userPrincipalName = $scriptState.userName

		# Get assignments for this script
		$assignments = $scriptInfo.assignments
		$anyAssignmentFound = $false
		
		if ($assignments -and $assignments.Count -gt 0) {
			foreach ($assignment in $assignments) {
				$thisAssignmentMatches = $false
				$context = '_unknown'
				$assignmentGroup = $null
				$assignmentGroupId = $null
				$groupType = ''
				$groupMembers = 'N/A'
				$assignmentGroupTooltip = ''
				$filterName = ''
				$filterMode = ''
				$filterTooltip = ''
				
				# Get filter information
				$filterId = $assignment.target.deviceAndAppManagementAssignmentFilterId
				$filterType = $assignment.target.deviceAndAppManagementAssignmentFilterType
				
				if ($filterType -and $filterType -ne 'none') {
					$filterMode = $filterType
					$filterObj = Get-ObjectFromGUID -Id $filterId
					if ($filterObj) {
						$filterName = $filterObj.displayName
						$filterTooltip = $filterObj.rule
					} else {
						$filterName = $filterId
					}
				}

				# Check assignment type
				if ($assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
					$assignmentGroup = 'All Users'
					$context = 'User'
					$assignmentGroupTooltip = 'Built-in All Users group'
					$groupMembers = ''
					$thisAssignmentMatches = $true
				}
				elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
					$assignmentGroup = 'All Devices'
					$context = 'Device'
					$assignmentGroupTooltip = 'Built-in All Devices group'
					$groupMembers = ''
					$thisAssignmentMatches = $true
				}
				elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget') {
					$groupId = $assignment.target.groupId
					
					# Check if device is member of this group
					$deviceGroupObj = $Script:deviceGroupMemberships | Where-Object { $_.id -eq $groupId }
					if ($deviceGroupObj) {
						$assignmentGroup = $deviceGroupObj.displayName
						$assignmentGroupId = $groupId
						$context = 'Device'
						$groupType = $deviceGroupObj.YodamiittiCustomMembershipType
						$assignmentGroupTooltip = $deviceGroupObj.membershipRule
						
						$devCount = $deviceGroupObj.YodamiittiCustomGroupMembersCountDevices
						$userCount = $deviceGroupObj.YodamiittiCustomGroupMembersCountUsers
						$groupMembers = ''
						if ($devCount -gt 0) { $groupMembers += "$devCount devices " }
						if ($userCount -gt 0) { $groupMembers += "$userCount users " }
						
						$thisAssignmentMatches = $true
					}
					
					# Check if primary user is member of this group
					if ($Script:PrimaryUserGroupsMemberOf | Where-Object { $_.id -eq $groupId }) {
						if ($assignmentGroup) {
							$context = '_Device/User'
						} else {
							$primaryGroupObj = $Script:PrimaryUserGroupsMemberOf | Where-Object { $_.id -eq $groupId }
							$assignmentGroup = $primaryGroupObj.displayName
							$assignmentGroupId = $groupId
							$context = 'User'
							$groupType = $primaryGroupObj.YodamiittiCustomMembershipType
							$assignmentGroupTooltip = $primaryGroupObj.membershipRule
							
							$devCount = $primaryGroupObj.YodamiittiCustomGroupMembersCountDevices
							$userCount = $primaryGroupObj.YodamiittiCustomGroupMembersCountUsers
							$groupMembers = ''
							if ($devCount -gt 0) { $groupMembers += "$devCount devices " }
							if ($userCount -gt 0) { $groupMembers += "$userCount users " }
						}
						$thisAssignmentMatches = $true
					}
					
					# Check if latest user is member of this group (if different from primary)
					if ($Script:PrimaryUser.id -ne $Script:LatestCheckedinUser.id) {
						if ($Script:LatestCheckedInUserGroupsMemberOf | Where-Object { $_.id -eq $groupId }) {
							if ($assignmentGroup) {
								if ($context -eq 'Device') {
									$context = '_Device/User'
								}
							} else {
								$latestGroupObj = $Script:LatestCheckedInUserGroupsMemberOf | Where-Object { $_.id -eq $groupId }
								$assignmentGroup = $latestGroupObj.displayName
								$assignmentGroupId = $groupId
								$context = 'User'
								$groupType = $latestGroupObj.YodamiittiCustomMembershipType
								$assignmentGroupTooltip = $latestGroupObj.membershipRule
								
								$devCount = $latestGroupObj.YodamiittiCustomGroupMembersCountDevices
								$userCount = $latestGroupObj.YodamiittiCustomGroupMembersCountUsers
								$groupMembers = ''
								if ($devCount -gt 0) { $groupMembers += "$devCount devices " }
								if ($userCount -gt 0) { $groupMembers += "$userCount users " }
							}
							$thisAssignmentMatches = $true
						}
					}
				}

				# Add schedule info to tooltip
				if ($assignment.runSchedule) {
					$scheduleType = $assignment.runSchedule.'@odata.type'
					$scheduleInterval = $assignment.runSchedule.interval
					$scheduleTime = $assignment.runSchedule.time
					
					if ($scheduleType -eq '#microsoft.graph.deviceHealthScriptRunOnceSchedule') {
						$scheduleDate = $assignment.runSchedule.date
						$assignmentGroupTooltip += "`n`nRemediation schedule:`nRun once`n$scheduleTime`n$scheduleDate"
					}
					elseif ($scheduleType -eq '#microsoft.graph.deviceHealthScriptHourlySchedule') {
						$assignmentGroupTooltip += "`n`nRemediation schedule:`nRun every $scheduleInterval hours"
					}
					elseif ($scheduleType -eq '#microsoft.graph.deviceHealthScriptDailySchedule') {
						$assignmentGroupTooltip += "`n`nRemediation schedule:`nRun every $scheduleInterval days"
					}
					else {
						$assignmentGroupTooltip += "`n`nRemediation schedule:`n$scheduleType`n$scheduleInterval`n$scheduleTime"
					}
				}

				# Only add if this specific assignment matches the device/user
				if ($thisAssignmentMatches) {
					$anyAssignmentFound = $true
					
					# Track remediation scripts for extended report download
					if ($ExtendedReport -and $scriptState.policyId) {
						if ($script:ScriptIdsToDownload -notcontains $scriptState.policyId) {
							Write-Verbose "Tracking remediation script ID: $($scriptState.policyId)"
							$script:ScriptIdsToDownload += $scriptState.policyId
						}
					}
					
					$results += [PSCustomObject]@{
						id = $scriptState.policyId
						context = $context
						scriptType = 'Remediation'
						displayName = if ($scriptInfo) { $scriptInfo.displayName } else { $scriptState.policyId }
						detectionStatus = $detectionStatus
						detectionStatusTooltip = $detectionTooltip
						remediationStatus = $remediationStatus
						remediationStatusTooltip = $remediationTooltip
						userPrincipalName = $userPrincipalName
						statusUpdateTime = $statusUpdateTime
						statusUpdateTimeTooltip = $statusUpdateTimeTooltip
						groupType = $groupType
						assignmentGroup = $assignmentGroup
						assignmentGroupTooltip = $assignmentGroupTooltip
						groupMembers = $groupMembers
						filter = $filterName
						filterMode = $filterMode
						filterTooltip = $filterTooltip
					}
				}
			}
		}
		
		# If no assignments matched, add entry without assignment info
		if (-not $anyAssignmentFound) {
			# Track remediation scripts for extended report download
			if ($ExtendedReport -and $scriptState.policyId) {
				if ($script:ScriptIdsToDownload -notcontains $scriptState.policyId) {
					Write-Verbose "Tracking remediation script ID (no assignment): $($scriptState.policyId)"
					$script:ScriptIdsToDownload += $scriptState.policyId
				}
			}
			
			$results += [PSCustomObject]@{
				id = $scriptState.policyId
				context = ''
				scriptType = 'Remediation'
				displayName = if ($scriptInfo) { $scriptInfo.displayName } else { $scriptState.policyId }
				detectionStatus = $detectionStatus
				detectionStatusTooltip = $detectionTooltip
				remediationStatus = $remediationStatus
				remediationStatusTooltip = $remediationTooltip
				userPrincipalName = $userPrincipalName
				statusUpdateTime = $statusUpdateTime
				statusUpdateTimeTooltip = $statusUpdateTimeTooltip
				groupType = ''
				assignmentGroup = 'No assignments'
				assignmentGroupTooltip = ''
				groupMembers = ''
				filter = ''
				filterMode = ''
				filterTooltip = ''
			}
		}
	}

	# Process platform scripts assignments (no device-specific state, just assignments)
	foreach ($platformScript in $script:PlatformScriptsWithAssignments) {
		$anyAssignmentFound = $false
		$assignments = $platformScript.assignments

		if ($assignments -and $assignments.Count -gt 0) {
			foreach ($assignment in $assignments) {
				$thisAssignmentMatches = $false
				$context = '_unknown'
				$assignmentGroup = $null
				$assignmentGroupId = $null
				$groupType = ''
				$groupMembers = 'N/A'
				$assignmentGroupTooltip = ''
				$filterName = ''
				$filterMode = ''
				$filterTooltip = ''
				
				# Get filter information
				$filterId = $assignment.target.deviceAndAppManagementAssignmentFilterId
				$filterType = $assignment.target.deviceAndAppManagementAssignmentFilterType
				
				if ($filterType -and $filterType -ne 'none') {
					$filterMode = $filterType
					$filterObj = Get-ObjectFromGUID -Id $filterId
					if ($filterObj) {
						$filterName = $filterObj.displayName
						$filterTooltip = $filterObj.rule
					} else {
						$filterName = $filterId
					}
				}

				# Check assignment type
				if ($assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
					$assignmentGroup = 'All Users'
					$context = 'User'
					$assignmentGroupTooltip = 'Built-in All Users group'
					$groupMembers = ''
					$thisAssignmentMatches = $true
				}
				elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
					$assignmentGroup = 'All Devices'
					$context = 'Device'
					$assignmentGroupTooltip = 'Built-in All Devices group'
					$groupMembers = ''
					$thisAssignmentMatches = $true
				}
				elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget') {
					$groupId = $assignment.target.groupId
					
					# Check if device is member of this group
					$deviceGroupObj = $Script:deviceGroupMemberships | Where-Object { $_.id -eq $groupId }
					if ($deviceGroupObj) {
						$assignmentGroup = $deviceGroupObj.displayName
						$assignmentGroupId = $groupId
						$context = 'Device'
						$groupType = $deviceGroupObj.YodamiittiCustomMembershipType
						$assignmentGroupTooltip = $deviceGroupObj.membershipRule
						
						$devCount = $deviceGroupObj.YodamiittiCustomGroupMembersCountDevices
						$userCount = $deviceGroupObj.YodamiittiCustomGroupMembersCountUsers
						$groupMembers = ''
						if ($devCount -gt 0) { $groupMembers += "$devCount devices " }
						if ($userCount -gt 0) { $groupMembers += "$userCount users " }
						
						$thisAssignmentMatches = $true
					}
					
					# Check if primary user is member of this group
					if ($Script:PrimaryUserGroupsMemberOf | Where-Object { $_.id -eq $groupId }) {
						if ($assignmentGroup) {
							$context = '_Device/User'
						} else {
							$primaryGroupObj = $Script:PrimaryUserGroupsMemberOf | Where-Object { $_.id -eq $groupId }
							$assignmentGroup = $primaryGroupObj.displayName
							$assignmentGroupId = $groupId
							$context = 'User'
							$groupType = $primaryGroupObj.YodamiittiCustomMembershipType
							$assignmentGroupTooltip = $primaryGroupObj.membershipRule
							
							$devCount = $primaryGroupObj.YodamiittiCustomGroupMembersCountDevices
							$userCount = $primaryGroupObj.YodamiittiCustomGroupMembersCountUsers
							$groupMembers = ''
							if ($devCount -gt 0) { $groupMembers += "$devCount devices " }
							if ($userCount -gt 0) { $groupMembers += "$userCount users " }
						}
						$thisAssignmentMatches = $true
					}
					
					# Check if latest user is member of this group (if different from primary)
					if ($Script:PrimaryUser.id -ne $Script:LatestCheckedinUser.id) {
						if ($Script:LatestCheckedInUserGroupsMemberOf | Where-Object { $_.id -eq $groupId }) {
							if ($assignmentGroup) {
								if ($context -eq 'Device') {
									$context = '_Device/User'
								}
							} else {
								$latestGroupObj = $Script:LatestCheckedInUserGroupsMemberOf | Where-Object { $_.id -eq $groupId }
								$assignmentGroup = $latestGroupObj.displayName
								$assignmentGroupId = $groupId
								$context = 'User'
								$groupType = $latestGroupObj.YodamiittiCustomMembershipType
								$assignmentGroupTooltip = $latestGroupObj.membershipRule
								
								$devCount = $latestGroupObj.YodamiittiCustomGroupMembersCountDevices
								$userCount = $latestGroupObj.YodamiittiCustomGroupMembersCountUsers
								$groupMembers = ''
								if ($devCount -gt 0) { $groupMembers += "$devCount devices " }
								if ($userCount -gt 0) { $groupMembers += "$userCount users " }
							}
							$thisAssignmentMatches = $true
						}
					}
				}

				# Only add if this specific assignment matches the device/user
				if ($thisAssignmentMatches) {
					# Check if script platform matches device OS
					# Skip scripts that don't match the device OS (e.g., don't show macOS scripts on Windows devices)
					# This applies to ALL assignments (device and user) because Intune filters by OS at deployment
					$scriptPlatform = $platformScript.ScriptPlatform
					$deviceOS = $script:IntuneManagedDevice.operatingSystem
					
					$platformMatches = $true
					# Check OS compatibility for all assignment types
					if ($scriptPlatform -eq 'Windows' -and $deviceOS -notlike 'Windows*') {
						$platformMatches = $false
					}
					elseif ($scriptPlatform -eq 'macOS' -and $deviceOS -ne 'macOS') {
						$platformMatches = $false
					}
					elseif ($scriptPlatform -eq 'Linux' -and $deviceOS -ne 'Linux') {
						$platformMatches = $false
					}
					
					if (-not $platformMatches) {
						Write-Verbose "Skipping $scriptPlatform script '$($platformScript.displayName)' - doesn't match device OS: $deviceOS"
						continue
					}
					
					$anyAssignmentFound = $true
					
					# Determine script type based on platform
					$scriptType = "Platform ($($platformScript.ScriptPlatform))"
					
					# Get display name (different property for Linux scripts)
					$displayName = if ($platformScript.displayName) { 
						$platformScript.displayName 
					} elseif ($platformScript.name) { 
						$platformScript.name 
					} else { 
						$platformScript.id 
					}

					# Track Windows PowerShell scripts and macOS shell scripts for extended report download
					if ($ExtendedReport -and ($platformScript.ScriptPlatform -eq 'Windows' -or $platformScript.ScriptPlatform -eq 'macOS')) {
						if ($platformScript.id -and $script:ScriptIdsToDownload -notcontains $platformScript.id) {
							Write-Verbose "Tracking platform script ID: $($platformScript.id) - $($displayName) (Platform: $($platformScript.ScriptPlatform))"
							$script:ScriptIdsToDownload += $platformScript.id
						}
					}

					$results += [PSCustomObject]@{
						id = $platformScript.id
						context = $context
						scriptType = $scriptType
						displayName = $displayName
						detectionStatus = 'N/A'
						detectionStatusTooltip = 'Platform scripts do not have detection state'
						remediationStatus = 'N/A'
						remediationStatusTooltip = 'Platform scripts do not have remediation state'
						userPrincipalName = ''
						statusUpdateTime = ''
						statusUpdateTimeTooltip = ''
						groupType = $groupType
						assignmentGroup = $assignmentGroup
						assignmentGroupTooltip = $assignmentGroupTooltip
						groupMembers = $groupMembers
						filter = $filterName
						filterMode = $filterMode
						filterTooltip = $filterTooltip
					}
				}
			}
		}
	}

	return $results
}


function New-IntuneDeviceHtmlReport {
	param(
		[hashtable]$Context
	)
	$device = $Context.ManagedDevice
	$azureDevice = $Context.AzureDevice
	$primaryContext = $Context.PrimaryUser
	$primaryUser = if ($primaryContext) { $primaryContext.User } else { $null }
	$primaryUserGroups = if ($primaryContext -and $primaryContext.Groups) { $primaryContext.Groups } else { @() }
	$latestContext = $Context.LatestUser
	$latestUser = if ($latestContext) { $latestContext.LatestUser } else { $null }
	$latestUserGroups = if ($latestContext -and $latestContext.LatestGroups) { $latestContext.LatestGroups } else { @() }
	$appAssignments = if ($Context.AppAssignments) { $Context.AppAssignments.Items } else { @() }
	$configPolicies = $Context.ConfigurationPolicies
	$deviceGroups = $Context.DeviceGroups
	$autopilotContext = $Context.Autopilot
	$autopilotDetail = if ($autopilotContext) { $autopilotContext.Detail } else { $null }
	$autopilotDeviceInfo = if ($autopilotContext) { $autopilotContext.Device } else { $null }
	$espContext = $Context.EnrollmentStatusPage

	function ConvertTo-FriendlyBytes {
		param($Bytes)

		if ($null -eq $Bytes) { return $null }
		try { $value = [double]$Bytes } catch { return $null }
		$units = @('B','KB','MB','GB','TB','PB')
		$index = 0
		while ($value -ge 1024 -and $index -lt $units.Count - 1) {
			$value /= 1024
			$index++
		}
		return ('{0:N0} {1}' -f $value, $units[$index])
	}

	function Get-StorageSummary {
		param($TotalBytes,$FreeBytes)

		try { $total = if ($null -ne $TotalBytes) { [double]$TotalBytes } else { $null } } catch { $total = $null }
		try { $free = if ($null -ne $FreeBytes) { [double]$FreeBytes } else { $null } } catch { $free = $null }
		if ($total -and $free) {
			return ('{0} / {1}' -f (ConvertTo-FriendlyBytes $free), (ConvertTo-FriendlyBytes $total))
		}
		elseif ($total) {
			return ConvertTo-FriendlyBytes $total
		}
		return $null
	}

	function New-DeviceDetailCards {
		param([array]$Items)

		if (-not $Items -or $Items.Count -eq 0) { return '' }
		$encode = {
			param($value)
			if ($null -eq $value) { return 'n/a' }
			$text = [string]$value
			if ([string]::IsNullOrWhiteSpace($text)) { $text = 'n/a' }
			return [System.Net.WebUtility]::HtmlEncode($text)
		}
		$cards = foreach ($item in $Items) {
			$label = & $encode $item.Label
			$value = & $encode $item.Value
			$tooltipAttr = ''
			if ($item.Tooltip) {
				$tooltipAttr = " data-tooltip=`"$(& $encode $item.Tooltip)`""
			}
			$accentClass = if ($item.Accent) { " $($item.Accent)" } else { '' }
			$secondary = if ($item.Secondary) { "<div class='info-secondary'>$(& $encode $item.Secondary)</div>" } else { '' }
			"<div class='info-card$accentClass'$tooltipAttr><div class='info-label'>$label</div><div class='info-value'>$value</div>$secondary</div>"
		}
		return "<div class='info-grid'>$([string]::Join('', $cards))</div>"
	}
	$now = Get-Date
	$css = @"
body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; background: #f5f7fb; color: #1f2933; }
.page { padding: 10px; }
.card { background: #fff; border-radius: 10px; box-shadow: 0 10px 25px rgba(15,23,42,.12); padding: 12px; margin-bottom: 10px; }
.grid { display: grid; gap: 16px; }
.grid-2 { grid-template-columns: repeat(auto-fit,minmax(220px,1fr)); }
.title { font-size: 32px; font-weight: 700; margin-bottom: 4px; }
.subtitle { color: #64748b; margin-bottom: 24px; }
.card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
h3 { margin: 0; font-weight: 700; font-size: 18px; }
.card-controls { display: flex; gap: 4px; }
.card-control-btn { background: #f1f5f9; border: 1px solid #cbd5e1; border-radius: 4px; padding: 4px 8px; cursor: pointer; font-size: 11px; font-weight: 600; color: #475569; transition: all 0.2s; }
.card-control-btn:hover { background: #e2e8f0; color: #1e293b; }
.card-control-btn.active { background: #3b82f6; color: #fff; border-color: #3b82f6; }
.card.minimized .card-body { display: none; }
.card.fullsize { position: relative; z-index: 100; }
.card.fullsize .card-body { display: block; }
.card.fullsize .card-body > table, .card.fullsize .card-body > .tab-content { flex: 1; }
.card.fullsize table { height: 100%; }
.info-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:8px; margin-top:10px; }
.info-card { background:#f8fafc; border-radius:10px; padding:8px 10px; border:1px solid #e2e8f0; position:relative; box-shadow:0 4px 12px rgba(15,23,42,.08); transition:all .2s ease; z-index:1; }
.info-card:hover { transform:translateY(-2px); box-shadow:0 12px 28px rgba(15,23,42,.18); z-index:2; }
.info-card[data-tooltip] { cursor: help; }
.tooltip-box { position: fixed; padding: 12px 16px; background: #f8fafc; color: #1e293b; border-radius: 6px; border: 1px solid #cbd5e1; font-family: 'Consolas','SFMono-Regular',monospace; font-size: 11px; white-space: pre-wrap; word-wrap: break-word; z-index: 999999; box-shadow: 0 8px 24px rgba(0,0,0,0.2); max-height: 80vh; max-width: min(600px, calc(100vw - 20px)); overflow: auto; line-height: 1.4; display: none; user-select: text; }
.info-label { font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.08em; color:#64748b; margin-bottom:4px; display:flex; align-items:center; gap:6px; }
.info-value { font-size:14px; font-weight:600; color:#0f172a; word-break:break-word; line-height:1.15; }
.info-secondary { margin-top:3px; font-size:10.5px; color:#475569; }
.info-card.accent-hardware { border-top:4px solid #3b82f6; }
.info-card.accent-network { border-top:4px solid #0ea5e9; }
.info-card.accent-status { border-top:4px solid #22c55e; }
.info-card.accent-security { border-top:4px solid #f97316; }
.info-card.accent-autopilot { border-top:4px solid #a855f7; }
.info-card.accent-user { border-top:4px solid #06b6d4; }
.info-card.accent-warning { border-top:4px solid #eab308; background:#fefce8; }
.info-card.accent-error { border-top:4px solid #ef4444; background:#fef2f2; }
.badge { display: inline-block; padding: 4px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; }
.badge-green { background:#d1fae5; color:#065f46; }
.badge-red { background:#fee2e2; color:#991b1b; }
.badge-yellow { background:#fef3c7; color:#92400e; }
.badge-column { display: flex; flex-direction: column; gap: 8px; }
.device-info-layout { display: grid; grid-template-columns: auto 1fr; gap: 20px; align-items: start; }
table { width: 100%; border-collapse: collapse; margin-top: 8px; }
th,td { padding: 4px 6px; border-bottom: 1px solid #e2e8f0; text-align: left; }
th { background:#f8fafc; font-size: 12px; text-transform: uppercase; letter-spacing: .06em; color:#475569; }
.apps-table,.config-table,.data-table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 12.5px; }
.apps-table th,.config-table th,.data-table th { background:#1e293b; color:#f8fafc; font-weight: 600; border-bottom: none; }
.apps-table td,.config-table td,.data-table td { background:#fff; border-bottom: 1px solid #e2e8f0; vertical-align: middle; }
.apps-table tr:hover td,.config-table tr:hover td,.data-table tr:hover td { background:#f1f5f9; }
.apps-table td.center,.config-table td.center,.data-table td.center { text-align: center; }
.apps-table td.right,.config-table td.right,.data-table td.right { text-align: right; }
.apps-table td.col-display strong,.config-table td.col-display strong,.data-table td.col-display strong { font-weight: 700; }
.apps-table td.warning-cell,.config-table td.warning-cell,.data-table td.warning-cell { background:#fef3c7; }
.apps-table td.danger-cell,.config-table td.danger-cell,.data-table td.danger-cell { background:#fecaca; }
.apps-table td.success-cell,.config-table td.success-cell,.data-table td.success-cell { background:#d1fae5; }
.apps-table td.shrink,.config-table td.shrink,.data-table td.shrink { width: 80px; }
.apps-table td.col-odatatype,.config-table td.col-odatatype,.data-table td.col-odatatype { width: 180px; font-family: 'Consolas','SFMono-Regular',monospace; font-size: 12px; }
.apps-table td.col-display,.config-table td.col-display,.data-table td.col-display { width: 280px; }
.apps-table td.col-group,.config-table td.col-group,.data-table td.col-group { width: 260px; }
.apps-table td.col-filter,.config-table td.col-filter,.data-table td.col-filter { width: 220px; }
.apps-table td.col-filterMode,.config-table td.col-filterMode,.data-table td.col-filterMode { width: 120px; }
.apps-table td.col-context,.config-table td.col-context,.data-table td.col-context { width: 90px; }
.apps-table td.col-members,.config-table td.col-members,.data-table td.col-members { width: 120px; }
.apps-table td.col-members,.config-table td.col-members,.data-table td.col-members { width: 110px; }
.group-membership-table tr.role-directory td { background: #fef9c3; }
.group-membership-table tr.role-globaladmin td { background: #fee2e2; font-weight: 600; }
.group-membership-table td.col-rule { font-family: 'Consolas','SFMono-Regular',monospace; font-size: 12px; }
.group-membership-table td.center { text-align: center; }
.config-table td.col-upn { width: 220px; font-family: 'Consolas','SFMono-Regular',monospace; font-size: 12px; }
.config-table td.col-state { width: 120px; }
.config-table td.col-groupType { width: 110px; }
.table-search { display:flex; align-items:center; gap:6px; margin-top:0; }
.table-search label { font-weight:600; color:#334155; font-size:12px; }
.table-search input { width:300px; padding:4px 8px; border-radius:6px; border:1px solid #cbd5e1; font-size:12px; }
.table-search .clear-btn { background:#ef4444; color:#fff; border:none; padding:4px 10px; border-radius:6px; cursor:pointer; font-size:12px; font-weight:600; transition:background 0.2s; }
.table-search .clear-btn:hover { background:#dc2626; }
.sortable th { cursor:pointer; position:relative; padding-right:18px; }
.sortable th::after { content:'\25B4\25BE'; position:absolute; right:6px; top:50%; transform:translateY(-50%); font-size:11px; color:#94a3b8; }
.sortable th[data-sort-dir="asc"]::after { content:'\25B4'; color:#38bdf8; }
.sortable th[data-sort-dir="desc"]::after { content:'\25BE'; color:#38bdf8; }
.table-scroll-wrapper { max-height: 25vh; overflow-y: auto; overflow-x: auto; border: 1px solid #e2e8f0; border-radius: 6px; margin-top: 8px; position: relative; z-index: 1; }
.card.fullsize .table-scroll-wrapper { max-height: 94vh; height: auto; }
.table-scroll-wrapper table { margin-top: 0; border-radius: 0; }
.table-scroll-wrapper thead th { position: sticky; top: 0; z-index: 1; box-shadow: 0 2px 2px -1px rgba(0,0,0,0.1); }
.tabs { display: flex; gap: 4px; border-bottom: 2px solid #e2e8f0; margin-bottom: 12px; }
.tab-button { background: #f8fafc; border: none; padding: 8px 16px; cursor: pointer; font-size: 13px; font-weight: 600; color: #64748b; border-radius: 6px 6px 0 0; transition: all 0.2s; }
.tab-button:hover { background: #e2e8f0; color: #334155; }
.tab-button.active { background: #3b82f6; color: #fff; }
.tab-content { display: none; }
.tab-content.active { display: block; }
details { margin-top: 10px; }
summary { cursor: pointer; font-weight: 600; }
pre { background:#f8fafc; color:#334155; padding:12px; border-radius:10px; overflow:auto; font-size: 11.5px; border: 1px solid #e2e8f0; }
.report-header { display: flex; justify-content: space-between; align-items: center; padding: 6px 20px; background: #f8fafc; border-bottom: 2px solid #e2e8f0; margin-bottom: 3px; font-size: 13px; color: #475569; }
.report-header .left { flex: 1; text-align: left; font-weight: 700; }
.report-header .center { flex: 1; text-align: center; font-weight: 700; font-size: 18px; color: #0f172a; }
.report-header .right { flex: 1; text-align: right; font-weight: 700; }
.report-footer { margin-top: 3px; padding: 10px 20px; background: #f8fafc; border-top: 2px solid #e2e8f0; color: #475569; display: flex; align-items: center; justify-content: center; font-size: 12px; }
.report-footer .creator-info { display: flex; flex-direction: row; align-items: center; margin-right: 15px; }
.report-footer .creator-info p { line-height: 1.2; margin: 0; }
.report-footer .creator-info p.author-text { margin-right: 15px; }
.profile-container { position: relative; width: 50px; height: 50px; border-radius: 50%; overflow: hidden; margin-right: 10px; }
.profile-container img { width: 100%; height: 100%; object-fit: cover; transition: opacity 0.3s; }
.profile-container img.black-profile { position: absolute; top: 0; left: 0; z-index: 1; }
.profile-container:hover img.black-profile { opacity: 0; }
.report-footer .company-logo { width: 100px; height: auto; margin: 0 20px; }
.report-footer a { color: #0066cc; text-decoration: none; font-weight: 600; }
.report-footer a:hover { text-decoration: underline; }
.row-details-modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 100000; align-items: center; justify-content: center; }
.row-details-modal.active { display: flex; }
.row-details-content { background: white; border-radius: 12px; width: 90%; min-width: 600px; max-width: 800px; max-height: 90vh; overflow: auto; padding: 24px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); position: relative; }
.row-details-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 2px solid #e2e8f0; padding-bottom: 12px; }
.row-details-header h2 { margin: 0; color: #1e293b; font-size: 20px; }
.row-details-close { background: #ef4444; color: white; border: none; border-radius: 6px; padding: 8px 16px; cursor: pointer; font-weight: 600; font-size: 14px; transition: background 0.2s; }
.row-details-close:hover { background: #dc2626; }
.row-details-body { display: grid; gap: 16px; }
.row-detail-item { background: #f8fafc; padding: 12px; border-radius: 8px; border-left: 4px solid #3b82f6; }
.row-detail-label { font-weight: 700; font-size: 12px; text-transform: uppercase; color: #64748b; margin-bottom: 6px; letter-spacing: 0.05em; }
.row-detail-value { color: #1e293b; font-size: 14px; word-break: break-word; white-space: pre-wrap; }
.row-detail-tooltip { margin-top: 8px; padding-top: 8px; border-top: 1px solid #cbd5e1; }
.row-detail-tooltip-label { font-weight: 700; font-size: 11px; text-transform: uppercase; color: #94a3b8; margin-bottom: 4px; }
.row-detail-tooltip-value { color: #475569; font-size: 13px; font-family: 'Consolas','SFMono-Regular',monospace; white-space: pre-wrap; }
tbody tr { cursor: pointer; }
tbody tr:hover { background-color: #f1f5f9; }
@media print { .card { box-shadow: none; } }
"@
	$complianceBadge = switch ($device.complianceState) {
		'compliant' { '<span class="badge badge-green">Compliant</span>' }
		'noncompliant' { '<span class="badge badge-red">Non-compliant</span>' }
		Default { '<span class="badge badge-yellow">Unknown</span>' }
	}

	$autopilotBadge = if ($device.autopilotEnrolled) { 
		'<span class="badge badge-green">Autopilot</span>' 
	} elseif ($script:AutopilotDevicePreparationPolicyWithAssignments) { 
		'<span class="badge badge-green">Device Preparation</span>' 
	}

	$encryptionBadge = if ($device.isEncrypted -eq $true) { '<span class="badge badge-green">Encrypted</span>' } elseif ($device.isEncrypted -eq $false) { '<span class="badge badge-red">Not Encrypted</span>' } else { '<span class="badge badge-yellow">Unknown Encryption</span>' }

	$primaryUserHtml = if ($primaryUser) {
		"<strong>$($primaryUser.displayName)</strong><br/>$($primaryUser.userPrincipalName)<br/>$($primaryUser.jobTitle)" }
	else { 'Shared device' }
	$latestUserHtml = if ($latestUser) {
		"<strong>$($latestUser.displayName)</strong><br/>$($latestUser.userPrincipalName)" } else { 'n/a' }
	$hardwareInfo = $device.hardwareInformation
	$wifiMac = if ($device.wifiMacAddress) { $device.wifiMacAddress } elseif ($hardwareInfo.wifiMacAddress) { $hardwareInfo.wifiMacAddress } elseif ($hardwareInfo.wlanMacAddress) { $hardwareInfo.wlanMacAddress } else { $null }
	$ethernetMac = $device.ethernetMacAddress
	
	# Extract IP addresses
	$wifiIpAddress = if ($hardwareInfo.ipAddressV4) { $hardwareInfo.ipAddressV4 } else { $null }
	$ethernetIpAddresses = if ($hardwareInfo.wiredIPv4Addresses -and $hardwareInfo.wiredIPv4Addresses.Count -gt 0) { 
		$hardwareInfo.wiredIPv4Addresses -join ', '
	} else { $null }
	
	$storageSummary = Get-StorageSummary -TotalBytes $device.totalStorageSpaceInBytes -FreeBytes $device.freeStorageSpaceInBytes
	$storageTooltip = if ($device.totalStorageSpaceInBytes -or $device.freeStorageSpaceInBytes) {
		$freeFriendly = ConvertTo-FriendlyBytes $device.freeStorageSpaceInBytes
		$totalFriendly = ConvertTo-FriendlyBytes $device.totalStorageSpaceInBytes
		if ($freeFriendly -or $totalFriendly) { "Free: $freeFriendly | Total: $totalFriendly" } else { $null }
	} else { $null }
	$lastSyncLocal = ConvertTo-LocalDateTimeString $device.lastSyncDateTime
	$lastSyncRelative = if ($device.lastSyncDateTime) {
		$span = New-TimeSpan -Start $device.lastSyncDateTime -End $now
		if ($span.TotalDays -ge 1) { '{0:N1} days ago' -f $span.TotalDays }
		elseif ($span.TotalHours -ge 1) { '{0:N1} hours ago' -f $span.TotalHours }
		elseif ($span.TotalMinutes -ge 1) { '{0:N0} minutes ago' -f $span.TotalMinutes }
		else { 'moments ago' }
	} else { $null }
	$lastSyncTooltip = if ($device.lastSyncDateTime) { "UTC $($device.lastSyncDateTime.ToUniversalTime().ToString('yyyy-MM-dd HH:mm'))" } else { 'Device has not checked in since enrollment.' }
	$primaryUserSecondary = if ($primaryUser) { $primaryUser.userPrincipalName } else { 'No assigned user' }
		$primaryUserTooltip = if ($primaryUser) {
		$basicInfo = ($primaryUser | Select-Object -Property accountEnabled,displayName,userPrincipalName,mail,userType,mobilePhone,jobTitle,department,companyName,employeeId,employeeType,streetAddress,postalCode,state,country,officeLocation,usageLocation | Format-List | Out-String).Trim()
		$proxyAddresses = if ($primaryUser.proxyAddresses) { ($primaryUser.proxyAddresses | Out-String).Trim() } else { '' }
		$otherMails = if ($primaryUser.otherMails) { ($primaryUser.otherMails | Out-String).Trim() } else { '' }
		$onPremAttrs = ($primaryUser | Select-Object -Property onPremisesSamAccountName,onPremisesUserPrincipalName,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesDomainName,onPremisesDistinguishedName,onPremisesImmutableId | Format-List | Out-String).Trim()
		$onPremExtAttrs = if ($primaryUser.onPremisesExtensionAttributes) { ($primaryUser.onPremisesExtensionAttributes | Format-List | Out-String).Trim() } else { '' }
		
		$tooltip = "Basic info`r`n$basicInfo`r`n"
		if ($proxyAddresses) { $tooltip += "`r`nproxyAddresses`r`n$proxyAddresses`r`n" }
		if ($otherMails) { $tooltip += "`r`notherMails`r`n$otherMails`r`n" }
		$tooltip += "`r`nonPremisesAttributes`r`n$onPremAttrs"
		if ($onPremExtAttrs) { $tooltip += "`r`n`r`nonPremisesExtensionAttributes`r`n$onPremExtAttrs" }
		$tooltip
	} else { 'No user assigned as primary owner.' }
	$autopilotProfileName = $null
	$autopilotProfileSecondary = $null
	if ($autopilotDetail) {
		if ($autopilotDetail.deploymentProfile.displayName) {
			$autopilotProfileName = $autopilotDetail.deploymentProfile.displayName
		}
		elseif ($autopilotDetail.intendedDeploymentProfile.displayName) {
			$autopilotProfileName = $autopilotDetail.intendedDeploymentProfile.displayName
		}
		# Get deployment mode and join type for secondary text
		$profile = if ($autopilotDetail.DeploymentProfileDetail) { $autopilotDetail.DeploymentProfileDetail } elseif ($autopilotDetail.deploymentProfile) { $autopilotDetail.deploymentProfile } else { $null }
		if ($profile) {
			$deploymentMode = if ($profile.outOfBoxExperienceSettings.deviceUsageType -eq 'shared') { 'Self-Deploying' } else { 'User-Driven' }
			$joinType = if ($profile.hybridAzureADJoinSkipConnectivityCheck -eq $true) { 'Hybrid joined' } else { 'Entra joined' }
			$autopilotProfileSecondary = "$deploymentMode / $joinType"
		}
	}
	# Autopilot Device card values
	$autopilotGroupTag = if ($autopilotDeviceInfo) { "Grouptag:`n$($autopilotDeviceInfo.groupTag)" } else { $null }
	$autopilotDeviceTooltip = if ($autopilotDeviceInfo) {
		$settings = @"
User
$(if ($autopilotDeviceInfo.userPrincipalName) { $autopilotDeviceInfo.userPrincipalName } else { 'unassigned' })

Serial number
$($autopilotDeviceInfo.serialNumber)

Manufacturer
$($autopilotDeviceInfo.manufacturer)

Model
$($autopilotDeviceInfo.model)

Device name
$(if ($autopilotDeviceInfo.displayName) { $autopilotDeviceInfo.displayName } else { 'N/A' })

Group tag
$(if ($autopilotDeviceInfo.groupTag) { $autopilotDeviceInfo.groupTag } else { '' })

Profile status
$(
	$profileA = if ($autopilotDetail -and $autopilotDetail.deploymentProfile) { $autopilotDetail.deploymentProfile.displayName } else { $null }
	$profileB = if ($autopilotDetail -and $autopilotDetail.intendedDeploymentProfile) { $autopilotDetail.intendedDeploymentProfile.displayName } else { $null }

	if ([string]::IsNullOrWhiteSpace($profileA) -or [string]::IsNullOrWhiteSpace($profileB) -or ($profileA -ne $profileB)) {
		'Assigning'
	} else {
		'Assigned'
	}
)

Assigned profile
$(if ($autopilotDetail -and ($autopilotDetail.deploymentProfile.displayName -or $autopilotDetail.intendedDeploymentProfile.displayName)) { 
	if ($autopilotDetail.deploymentProfile.displayName) { $autopilotDetail.deploymentProfile.displayName } else { $autopilotDetail.intendedDeploymentProfile.displayName }
} else { '' })

Date assigned
$(if ($autopilotDeviceInfo.deploymentProfileAssignedDateTime) {
	([datetimeoffset]::Parse($autopilotDeviceInfo.deploymentProfileAssignedDateTime.ToString())).LocalDateTime.ToString([System.Globalization.CultureInfo]::CurrentCulture)
} else { '' })

Enrollment state
$(if ($autopilotDeviceInfo.enrollmentState) { $autopilotDeviceInfo.enrollmentState } else { '' })

Associated Intune device
$(if ($device.deviceName) { $device.deviceName } else { '' })

Associated Microsoft Entra device
$(if ($azureDevice.displayName) { $azureDevice.displayName } else { '' })

Last contacted
$(if ($autopilotDeviceInfo.lastContactedDateTime) {
	([datetimeoffset]::Parse($autopilotDeviceInfo.lastContactedDateTime.ToString())).LocalDateTime.ToString([System.Globalization.CultureInfo]::CurrentCulture)
} else { '' })

Purchase order
$(if ($autopilotDeviceInfo.purchaseOrderIdentifier) { $autopilotDeviceInfo.purchaseOrderIdentifier } else { 'N/A' })
"@
		[System.Net.WebUtility]::HtmlEncode($settings)
	} else { 'No Autopilot device information available.' }
	$autopilotProfileTooltip = if ($autopilotDetail) {
		$profile = if ($autopilotDetail.DeploymentProfileDetail) { $autopilotDetail.DeploymentProfileDetail } elseif ($autopilotDetail.deploymentProfile) { $autopilotDetail.deploymentProfile } else { $null }
		if ($profile) {
			$settings = @"
Name
$($profile.displayName)

Description
$(if ($profile.description) { $profile.description } else { 'No Description' })

Convert all targeted devices to Autopilot
$(if ($profile.hardwareHashExtractionEnabled) { 'Yes' } else { 'No' })

Device type
$(if ($profile.deviceType -eq 'windowsPc') { 'Windows PC' } else { $profile.deviceType })

Out-of-box experience (OOBE)

Deployment mode
$(if ($profile.outOfBoxExperienceSettings.deviceUsageType -eq 'singleUser') { 'User-Driven' } elseif ($profile.outOfBoxExperienceSettings.deviceUsageType -eq 'shared') { 'Self-Deploying' } else { $profile.outOfBoxExperienceSettings.deviceUsageType })

Join to Microsoft Entra ID as
$(if ($null -ne $profile.PSObject.Properties['hybridAzureADJoinSkipConnectivityCheck']) { 'Microsoft Entra hybrid joined' } else { 'Microsoft Entra joined' })

$(if ($null -ne $profile.PSObject.Properties['hybridAzureADJoinSkipConnectivityCheck']) { if ($profile.hybridAzureADJoinSkipConnectivityCheck) { "Skip AD connectivity check`nYes" } else { "Skip AD connectivity checkSkip AD connectivity check`nNo" } })

Microsoft Software License Terms
$(if ($profile.outOfBoxExperienceSettings.hideEULA) { 'Hide' } else { 'Show' })

Privacy settings
$(if ($profile.outOfBoxExperienceSettings.hidePrivacySettings) { 'Hide' } else { 'Show' })

Hide change account options
$(if ($profile.outOfBoxExperienceSettings.hideEscapeLink) { 'Hide' } else { 'Show' })

User account type
$(if ($profile.outOfBoxExperienceSettings.userType -eq 'administrator') { 'Administrator' } else { 'Standard' })

Allow pre-provisioned deployment
$(if ($profile.preprovisioningAllowed) { 'Yes' } else { 'No' })

Language (Region)
$(switch ([string]$profile.language) {
	'os-default' { 'Operating system default' }
	''           { 'User select' }
	Default      { "$($profile.language)" } # e.g. fi-FI
})

$(if ($profile.language -and ($profile.language -ne '')) { if ($null -ne $profile.outOfBoxExperienceSettings.PSObject.Properties['skipKeyboardSelectionPage'] -and $profile.outOfBoxExperienceSettings.skipKeyboardSelectionPage) { 'Automatically configure keyboard`nYes' } else { 'Automatically configure keyboard`nNo' } })

$(if ($profile.deviceNameTemplate) { "Apply device name template`nYes" } else { "Apply device name template`nNo" })

$(if ($profile.deviceNameTemplate) { "Enter a name`n$($profile.deviceNameTemplate)" })
"@
			if ($profile.assignments -and $profile.assignments.Count -gt 0) {
				$settings += "`r`n`r`nAssignments`r`n`r`nIncluded groups`r`n"
				foreach ($assignment in $profile.assignments) {
					if ($assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
						$settings += "   All Devices`r`n"
					}
					elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
						$settings += "   All Users`r`n"
					}
					elseif ($assignment.target.displayName) {
						$settings += "   $($assignment.target.displayName)`r`n"
					}
					elseif ($assignment.target.groupId) {
						$settings += "   $($assignment.target.groupId)`r`n"
					}
				}
			}
			[System.Net.WebUtility]::HtmlEncode($settings)
		}
		else {
			'No Autopilot profile information available.'
		}
	} else { 'No Autopilot profile information available.' }
	$espTooltip = if ($espContext -and $espContext.Detail) {
		$esp = $espContext.Detail
		$settings = @"
Name
$($esp.displayName)

Description
$($esp.description)

Settings

Show app and profile configuration progress
$(if ($esp.showInstallationProgress) { 'Yes' } else { 'No' })

Show an error when installation takes longer than specified number of minutes
$($esp.installProgressTimeoutInMinutes)

Show custom message when time limit or error occurs
$(if ($esp.allowDeviceUseOnInstallFailure) { 'Yes' } else { 'No' })

Error message
$($esp.customErrorMessage)

Turn on log collection and diagnostics page for end users
$(if ($esp.allowLogCollectionOnInstallFailure) { 'Yes' } else { 'No' })

Only show page to devices provisioned by out-of-box experience (OOBE)
$(if ($esp.disableUserStatusTrackingAfterFirstUser -and $esp.trackInstallProgressForAutopilotOnly) { 'Yes' } else { 'No' })

Install Windows updates (might restart the device)
$(if ($esp.installQualityUpdates) { 'Yes' } else { 'No' })

Allow users to reset device if installation error occurs
$(if ($esp.allowDeviceResetOnInstallFailure) { 'Yes' } else { 'No' })

Allow users to use device if installation error occurs
$(if ($esp.allowDeviceUseOnInstallFailure) { 'Yes' } else { 'No' })

Only fail selected blocking apps in technician phase
$(if ($esp.allowNonBlockingAppInstallation) { 'Yes' } else { 'No' })

Block device use until required apps are installed if they are assigned to the user/device
$(if ($esp.selectedMobileAppIds -and ($esp.selectedMobileAppIds.Count -gt 0)) { 'Selected' } else { 'All' })

$(if ($esp.selectedMobileAppIds -and ($esp.selectedMobileAppIds.Count -gt 0)) { (($esp.selectedMobileAppIds | ForEach-Object { "  $_" }) -join "`r`n") } else { '' })
"@
		if ($esp.assignments -and $esp.assignments.Count -gt 0) {
			$settings += "`r`n`r`nAssignments`r`n`r`nIncluded groups`r`n"
			foreach ($assignment in $esp.assignments) {
				if ($assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
					$settings += "   All Devices`r`n"
				}
				elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
					$settings += "   All Users`r`n"
				}
				elseif ($assignment.target.displayName) {
					$settings += "   $($assignment.target.displayName)`r`n"
				}
				elseif ($assignment.target.groupId) {
					$settings += "   $($assignment.target.groupId)`r`n"
				}
			}
		}
		[System.Net.WebUtility]::HtmlEncode($settings)
	} else { 'No Enrollment Status Page information available.' }

	# Autopilot Device Preparation tooltip
	$autopilotDevicePrepTooltip = if ($script:AutopilotDevicePreparationPolicyWithAssignments) {
		$policy = $script:AutopilotDevicePreparationPolicyWithAssignments
		
		# Parse settings for friendly display
		$deploymentMode = ''
		$deploymentType = ''
		$joinType = ''
		$accountType = ''
		$timeout = ''
		$customErrorMessage = ''
		$allowSkip = ''
		$allowDiagnostics = ''
		$allowedApps = @()
		$allowedScripts = @()
		
		foreach ($setting in $policy.settings) {
			$settingId = $setting.settingInstance.settingDefinitionId
			switch -Wildcard ($settingId) {
				'*_deploymentmode' {
					$value = $setting.settingInstance.choiceSettingValue.value
					$deploymentMode = if ($value -match '_0$') { 'User-driven' } elseif ($value -match '_1$') { 'Self-deploying' } else { $value }
				}
				'*_deploymenttype' {
					$value = $setting.settingInstance.choiceSettingValue.value
					$deploymentType = if ($value -match '_0$') { 'Single user' } elseif ($value -match '_1$') { 'Shared device' } else { $value }
				}
				'*_jointype' {
					$value = $setting.settingInstance.choiceSettingValue.value
					$joinType = if ($value -match '_0$') { 'Microsoft Entra joined' } elseif ($value -match '_1$') { 'Microsoft Entra hybrid joined' } else { $value }
				}
				'*_accountype' {
					$value = $setting.settingInstance.choiceSettingValue.value
					$accountType = if ($value -match '_0$') { 'Administrator' } elseif ($value -match '_1$') { 'Standard User' } else { $value }
				}
				'*_timeout' {
					$timeout = $setting.settingInstance.simpleSettingValue.value
				}
				'*_customerrormessage' {
					$customErrorMessage = $setting.settingInstance.simpleSettingValue.value
				}
				'*_allowskip' {
					$value = $setting.settingInstance.choiceSettingValue.value
					$allowSkip = if ($value -match '_0$') { 'No' } elseif ($value -match '_1$') { 'Yes' } else { $value }
				}
				'*_allowdiagnostics' {
					$value = $setting.settingInstance.choiceSettingValue.value
					$allowDiagnostics = if ($value -match '_0$') { 'No' } elseif ($value -match '_1$') { 'Yes' } else { $value }
				}
				'*_allowedappids' {
					foreach ($appValue in $setting.settingInstance.simpleSettingCollectionValue) {
						try {
							$appJson = $appValue.value | ConvertFrom-Json -ErrorAction Stop
							$appId = $appJson.id
							$appType = $appJson.type
							
							# Parse friendly app type name
							$friendlyType = switch -Wildcard ($appType) {
								'*win32LobApp' { 'Win32' }
								'*winGetApp' { 'WinGet' }
								'*officeSuiteApp' { 'Microsoft 365 Apps' }
								'*webApp' { 'Web App' }
								'*windowsMobileMSI' { 'MSI' }
								'*iosStoreApp' { 'iOS Store' }
								'*androidManagedStoreApp' { 'Android Store' }
								default { $appType -replace '#microsoft\.graph\.', '' }
							}
							
							$appName = Get-NameFromGUID -Id $appId
							if ($appName) {
								$allowedApps += "$appName ($friendlyType)"
							} else {
								$allowedApps += "$appId ($friendlyType)"
							}
						} catch {
							# If parsing fails, just add the GUID if it looks like one
							if ($appValue.value -match '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}') {
								$allowedApps += $Matches[0]
							} else {
								$allowedApps += $appValue.value
							}
						}
					}
				}
				'*_allowedscriptids' {
					foreach ($scriptValue in $setting.settingInstance.simpleSettingCollectionValue) {
						$scriptId = $scriptValue.value
						$scriptName = Get-NameFromGUID -Id $scriptId
						if ($scriptName) {
							$allowedScripts += $scriptName
						} else {
							$allowedScripts += $scriptId
						}
					}
				}
			}
		}
		
		$settings = @"
Name
$($policy.name)

Description
$(if ($policy.description) { $policy.description } else { '--' })

Deployment settings

Deployment mode
$deploymentMode

Deployment type
$deploymentType

Join type
$joinType

User account type
$accountType

Out-of-box experience settings

Minutes allowed before showing installation error
$timeout

Custom error message
$customErrorMessage

Allow users to skip setup after multiple attempts
$allowSkip

Show link to diagnostics
$allowDiagnostics
"@

		if ($allowedApps.Count -gt 0) {
			$settings += "`r`n`r`nApps`r`n`r`nAllowed Applications`r`n"
			foreach ($app in $allowedApps) {
				$settings += "  $app`r`n"
			}
		}
		
		if ($allowedScripts.Count -gt 0) {
			$settings += "`r`nScripts`r`n`r`nAllowed Scripts`r`n"
			foreach ($script in $allowedScripts) {
				$settings += "  $script`r`n"
			}
		}

		if ($policy.assignments -and $policy.assignments.Count -gt 0) {
			$settings += "`r`nAssignments`r`n`r`nIncluded groups`r`n"
			foreach ($assignment in $policy.assignments) {
				if ($assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
					$settings += "   All Devices`r`n"
				}
				elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
					$settings += "   All Users`r`n"
				}
				elseif ($assignment.target.displayName) {
					$settings += "   $($assignment.target.displayName)`r`n"
				}
				elseif ($assignment.target.groupId) {
					$groupName = Get-NameFromGUID -Id $assignment.target.groupId
					if ($groupName) {
						$settings += "   $groupName`r`n"
					} else {
						$settings += "   $($assignment.target.groupId)`r`n"
					}
				}
			}
		}
		[System.Net.WebUtility]::HtmlEncode($settings)
	} else { 'No Autopilot Device Preparation policy information available.' }

	$deviceNameTooltip = if ($device) {
		$basicDeviceInfo = ($device | Select-Object -Property userPrincipalName,operatingSystem,osVersion,ownerType,deviceType,Manufacturer,Model,chassisType,serialNumber,deviceEnrollmentType,joinType,managedDeviceName,autopilotEnrolled,enrollmentProfileName,enrolledDateTime | Format-List | Out-String).Trim()
		$azureExtAttrs = if ($azureDevice -and $azureDevice.extensionAttributes) { ($azureDevice.extensionAttributes | Format-List | Out-String).Trim() } else { '' }
		
		$tooltip = "Device Properties`r`n$basicDeviceInfo"
		if ($azureExtAttrs) { $tooltip += "`r`n`r`nEntraID device extensionAttributes`r`n$azureExtAttrs" }
		$tooltip
	} else { '' }
	$latestUserTooltip = if ($latestUser) {
		$basicInfo = ($latestUser | Select-Object -Property accountEnabled,displayName,userPrincipalName,mail,userType,mobilePhone,jobTitle,department,companyName,employeeId,employeeType,streetAddress,postalCode,state,country,officeLocation,usageLocation | Format-List | Out-String).Trim()
		$proxyAddresses = if ($latestUser.proxyAddresses) { ($latestUser.proxyAddresses | Out-String).Trim() } else { '' }
		$otherMails = if ($latestUser.otherMails) { ($latestUser.otherMails | Out-String).Trim() } else { '' }
		$onPremAttrs = ($latestUser | Select-Object -Property onPremisesSamAccountName,onPremisesUserPrincipalName,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesDomainName,onPremisesDistinguishedName,onPremisesImmutableId | Format-List | Out-String).Trim()
		$onPremExtAttrs = if ($latestUser.onPremisesExtensionAttributes) { ($latestUser.onPremisesExtensionAttributes | Format-List | Out-String).Trim() } else { '' }
		
		$tooltip = "Basic info`r`n$basicInfo`r`n"
		if ($proxyAddresses) { $tooltip += "`r`nproxyAddresses`r`n$proxyAddresses`r`n" }
		if ($otherMails) { $tooltip += "`r`notherMails`r`n$otherMails`r`n" }
		$tooltip += "`r`nonPremisesAttributes`r`n$onPremAttrs"
		if ($onPremExtAttrs) { $tooltip += "`r`n`r`nonPremisesExtensionAttributes`r`n$onPremExtAttrs" }
		$tooltip
	} else { 'No recent logon information available.' }
	$detailItems = @()
	
	# Settings Catalog Conflicts Card (ExtendedReport only)
	if ($ExtendedReport -and $script:SettingsCatalogConflicts -and $script:SettingsCatalogConflicts.HasIssues) {
		$conflictCount = $script:SettingsCatalogConflicts.Conflicts.Count
		$warningCount = $script:SettingsCatalogConflicts.Warnings.Count
		
		$conflictLabel = if ($conflictCount -gt 0) { "Settings Conflicts" } else { "Settings Warnings" }
		$conflictValue = if ($conflictCount -gt 0) { "$conflictCount conflicts detected" } else { "$warningCount same/duplicate settings" }
		$conflictAccent = if ($conflictCount -gt 0) { 'accent-error' } else { 'accent-warning' }
		$conflictSecondary = if ($conflictCount -gt 0 -and $warningCount -gt 0) { "Also $warningCount warnings" } elseif ($conflictCount -eq 0) { "Same value in multiple policies" } else { $null }
		
		# Build detailed tooltip
		$conflictTooltip = "Settings Catalog Conflict Analysis`n`n"
		
		if ($conflictCount -gt 0) {
			$conflictTooltip += "=== CONFLICTS (Different Values) ===`n"
			foreach ($conflict in $script:SettingsCatalogConflicts.Conflicts) {
				$conflictTooltip += "`nSetting: $($conflict.SettingName)`n"
				foreach ($instance in $conflict.Instances) {
					$conflictTooltip += "  • $($instance.PolicyName): $($instance.Value)`n"
				}
			}
			$conflictTooltip += "`n"
		}
		
		if ($warningCount -gt 0) {
			$conflictTooltip += "=== WARNINGS (Duplicate Settings) ===`n"
			foreach ($warning in $script:SettingsCatalogConflicts.Warnings) {
				$conflictTooltip += "`nSetting: $($warning.SettingName)`n"
				if ($warning.IsAdditive) {
					$conflictTooltip += "  Type: Additive (values merge across policies)`n"
					$conflictTooltip += "  Values:`n"
					foreach ($instance in $warning.Instances) {
						$conflictTooltip += "    • $($instance.PolicyName): $($instance.Value)`n"
					}
				} else {
					$conflictTooltip += "  Value: $($warning.Value)`n"
					$conflictTooltip += "  Configured in policies:`n"
					foreach ($instance in $warning.Instances) {
						$conflictTooltip += "    • $($instance.PolicyName)`n"
					}
				}
			}
		}
		
		$conflictTooltip += "`n"
		
		$detailItems += @{ Label=$conflictLabel; Value=$conflictValue; Secondary=$conflictSecondary; Accent=$conflictAccent; Tooltip=$conflictTooltip }
	}
	
	$detailItems += @{ Label='Computer name'; Value=$device.deviceName; Accent='accent-hardware'; Tooltip=$deviceNameTooltip }
	$detailItems += @{ Label='Primary user'; Value=$(if ($primaryUser) { $primaryUser.displayName } else { 'Unassigned' }); Secondary=$primaryUserSecondary; Accent='accent-user'; Tooltip=$primaryUserTooltip }
	$detailItems += @{ Label='Latest logon'; Value=$(if ($latestUser) { $latestUser.displayName } else { 'n/a' }); Secondary=$(if ($latestUser) { $latestUser.userPrincipalName } else { '' }); Accent='accent-user'; Tooltip=$latestUserTooltip }
	$detailItems += @{ Label='Manufacturer'; Value=$device.manufacturer; Accent='accent-hardware'; Tooltip='Hardware manufacturer reported by device inventory.' }
	$detailItems += @{ Label='Model'; Value=$device.model; Accent='accent-hardware'; Tooltip='Model reported by the managed device record.' }
	$detailItems += @{ Label='Serial'; Value=$device.serialNumber; Accent='accent-hardware'; Tooltip='Serial number synced from Intune hardware information.' }
	
	# OS build with Windows SKU if applicable
	$osBuildValue = "$($device.operatingSystem) $($device.osVersion)"
	$osBuildSecondary = if ($device.operatingSystem -like 'Windows*' -and $device.skuFamily) { $device.skuFamily } else { $null }
	$detailItems += @{ Label='OS build'; Value=$osBuildValue; Secondary=$osBuildSecondary; Accent='accent-hardware'; Tooltip='Operating system and version currently reported by the device.' }
	
	# OS Language (common for all operating systems)
	$osLanguage = if ($hardwareInfo -and $hardwareInfo.operatingSystemLanguage) { $hardwareInfo.operatingSystemLanguage } else { $null }
	$detailItems += @{ Label='OS Language'; Value=$osLanguage; Accent='accent-hardware'; Tooltip='Operating system language/locale reported by the device.' }
	
	$detailItems += @{ Label='Wi-Fi IP Address'; Value=$wifiIpAddress; Accent='accent-network'; Tooltip='Wi-Fi IPv4 address from hardware information.' }
	$detailItems += @{ Label='Ethernet IP Address'; Value=$ethernetIpAddresses; Accent='accent-network'; Tooltip='Wired IPv4 addresses from hardware information. Multiple addresses may be present.' }
	$detailItems += @{ Label='Wi-Fi MAC'; Value=$wifiMac; Accent='accent-network'; Tooltip='Wireless adapter MAC address from the managed device.' }
	$detailItems += @{ Label='Ethernet MAC'; Value=$ethernetMac; Accent='accent-network'; Tooltip='Primary wired MAC address returned by Graph.' }
	$detailItems += @{ Label='Storage Free/Total'; Value=$storageSummary; Accent='accent-hardware'; Tooltip=$storageTooltip }
	$detailItems += @{ Label='Ownership'; Value=$device.managedDeviceOwnerType; Accent='accent-status'; Tooltip='Ownership category (Corporate vs Personal).' }
	$detailItems += @{ Label='Last sync'; Value=$lastSyncLocal; Secondary=$lastSyncRelative; Accent='accent-status'; Tooltip=$lastSyncTooltip }
	
	# Enrollment type - map technical values to friendly names
	$enrollmentTypeValue = $device.deviceEnrollmentType
	$enrollmentTypeTooltip = 'Method used to enroll this device.'
	
	# Check for MDE (Microsoft Defender for Endpoint) managed devices first
	if ($device.managementAgent -eq 'msSense') {
		$enrollmentTypeValue = 'MDE'
		$enrollmentTypeTooltip = 'Microsoft Defender for Endpoint (MDE) managed device. Lightweight management through Microsoft Defender without full MDM enrollment.'
	}
	# Map technical enrollment type values to friendly names
	elseif ($enrollmentTypeValue) {
		$enrollmentTypeFriendlyNames = @{
			'unknown' = 'Unknown'
			'userEnrollment' = 'User Enrollment'
			'deviceEnrollmentManager' = 'Device Enrollment Manager (DEM)'
			'appleBulkWithUser' = 'Apple ADE with User Affinity'
			'appleBulkWithoutUser' = 'Apple ADE without User Affinity'
			'windowsAzureADJoin' = 'Entra Joined'
			'windowsBulkUserless' = 'Windows Bulk Userless'
			'windowsAutoEnrollment' = 'Windows Auto Enrollment (GPO)'
			'windowsBulkAzureDomainJoin' = 'Windows Bulk Entra Domain Join'
			'windowsCoManagement' = 'Co-Managed'
			'windowsAzureADJoinUsingDeviceAuth' = 'Entra Joined (Device Auth)'
			'appleUserEnrollment' = 'Apple User Enrollment'
			'appleUserEnrollmentWithServiceAccount' = 'Apple User Enrollment (Service Account)'
		}
		
		if ($enrollmentTypeFriendlyNames.ContainsKey($enrollmentTypeValue)) {
			$friendlyName = $enrollmentTypeFriendlyNames[$enrollmentTypeValue]
			$enrollmentTypeTooltip = "Enrollment Method: $friendlyName`n`nTechnical value: $enrollmentTypeValue"
			$enrollmentTypeValue = $friendlyName
		}
	}
	
	$detailItems += @{ Label='Enrollment type'; Value=$enrollmentTypeValue; Accent='accent-status'; Tooltip=$enrollmentTypeTooltip }
	
	# macOS/iOS-specific enrollment profile name
	if (($device.operatingSystem -like 'macOS*' -or $device.operatingSystem -like 'iOS*') -and $device.enrollmentProfileName) {
		$enrollmentProfileTooltip = 'The name of the enrollment profile used to enroll this device.'
		
		# If ExtendedReport is enabled and we have an enrollment profile name, fetch the profile details
		if ($script:AppleEnrollmentProfileDetails) {
			$enrollmentProfileTooltip = ConvertTo-ReadableAppleEnrollmentProfile -ProfileData $script:AppleEnrollmentProfileDetails
			if (-not $enrollmentProfileTooltip) {
				$enrollmentProfileTooltip = 'The name of the enrollment profile used to enroll this device.'
			}
		}
		
		$detailItems += @{ Label='Enrollment Profile Name'; Value=$device.enrollmentProfileName; Accent='accent-status'; Tooltip=$enrollmentProfileTooltip }
	}
	
	# Windows-only cards (Autopilot features are Windows-specific)
	if ($device.operatingSystem -like 'Windows*') {
		$detailItems += @{ Label='Autopilot profile'; Value=$autopilotProfileName; Secondary=$autopilotProfileSecondary; Accent='accent-autopilot'; Tooltip=$autopilotProfileTooltip }
		$detailItems += @{ Label='Autopilot Device'; Value=$autopilotGroupTag; Accent='accent-autopilot'; Tooltip=$autopilotDeviceTooltip }
		$detailItems += @{ Label='Enrollment Status Page'; Value=$(if ($espContext) { $espContext.Name } else { $null }); Accent='accent-autopilot'; Tooltip=$espTooltip }
		$detailItems += @{ Label='Autopilot Device Preparation'; Value=$(if ($script:AutopilotDevicePreparationPolicyWithAssignments) { $script:AutopilotDevicePreparationPolicyWithAssignments.name } else { $null }); Accent='accent-autopilot'; Tooltip=$autopilotDevicePrepTooltip }
		
		# Co-Management card (only if configurationManagerClientEnabledFeatures exists)
		if ($device.configurationManagerClientEnabledFeatures) {
			$coMgmtFeatures = $device.configurationManagerClientEnabledFeatures
			$coMgmtTooltip = @"
Configuration Manager Client Enabled Features

{0,-35} {1}
{2,-35} {3}
{4,-35} {5}
{6,-35} {7}
{8,-35} {9}
{10,-35} {11}
{12,-35} {13}
{14,-35} {15}
"@ -f 'Inventory:', $coMgmtFeatures.inventory,
      'Modern Apps:', $coMgmtFeatures.modernApps,
      'Resource Access:', $coMgmtFeatures.resourceAccess,
      'Device Configuration:', $coMgmtFeatures.deviceConfiguration,
      'Compliance Policy:', $coMgmtFeatures.compliancePolicy,
      'Windows Update for Business:', $coMgmtFeatures.windowsUpdateForBusiness,
      'Endpoint Protection:', $coMgmtFeatures.endpointProtection,
      'Office Apps:', $coMgmtFeatures.officeApps
			$detailItems += @{ Label='Co-Managed'; Value='Yes'; Accent='accent-status'; Tooltip=$coMgmtTooltip }
		}
	}
	
	$deviceInfoCards = New-DeviceDetailCards -Items $detailItems

	function ConvertTo-HtmlRows {
		param([array]$Items,[string[]]$Columns)

		$result = ''
		foreach ($row in $Items) {
			$cells = foreach ($col in $Columns) { "<td>$($row.$col)</td>" }
			$result += "<tr>$([string]::Join('', $cells))</tr>"
		}
		return $result
	}

	function New-PrimaryUserGroupMembershipTableHtml {
		param(
			[array]$Groups,
			[bool]$HasPrimaryUser
		)

		if (-not $HasPrimaryUser) {
			return '<p>No primary user assigned to this device, so no user group memberships are available.</p>'
		}
		if (-not $Groups -or $Groups.Count -eq 0) {
			return '<p>No primary user group memberships resolved.</p>'
		}

		$encode = {
			param($value)
			if ($null -eq $value) { return '' }
			return [System.Net.WebUtility]::HtmlEncode([string]$value)
		}

		$rows = foreach ($group in ($Groups | Sort-Object -Property displayName)) {
			$display = & $encode $group.displayName
			$devices = & $encode $group.YodamiittiCustomGroupMembersCountDevices
			$users = & $encode $group.YodamiittiCustomGroupMembersCountUsers
			$type = & $encode $group.YodamiittiCustomGroupType
			$security = & $encode $group.securityEnabled
			$membershipType = & $encode $group.YodamiittiCustomMembershipType
			$rule = & $encode $group.membershipRule
			$descriptionTooltip = if ([string]::IsNullOrWhiteSpace($group.description)) { '' } else { & $encode $group.description }
			$descriptionAttr = if ($descriptionTooltip) { " title='$descriptionTooltip'" } else { '' }
			$ruleTooltip = if ([string]::IsNullOrWhiteSpace($group.membershipRule)) { '' } else { & $encode $group.membershipRule }
			$ruleAttr = if ($ruleTooltip) { " title='$ruleTooltip'" } else { '' }

			$rowClass = ''
			if ($group.YodamiittiCustomGroupType -eq 'DirectoryRole') { $rowClass = 'role-directory' }
			if ($group.displayName -eq 'Global Administrator') { $rowClass = 'role-globaladmin' }
			$rowClassAttr = if ($rowClass) { " class='$rowClass'" } else { '' }

			"<tr$rowClassAttr><td class='col-display'$descriptionAttr><strong>$display</strong></td><td class='center'>$devices</td><td class='center'>$users</td><td class='col-groupType'>$type</td><td class='center'>$security</td><td class='col-groupType'>$membershipType</td><td class='col-rule'$ruleAttr>$rule</td></tr>"
		}

		$table = @()
		$table += '<div class="table-scroll-wrapper">'
		$table += '<table class="data-table sortable group-membership-table primary-group-table" data-table="Primary user group memberships">'
		$table += '  <thead>'
		$table += '    <tr>'
		$table += '      <th>Display name</th>'
		$table += '      <th>Devices</th>'
		$table += '      <th>Users</th>'
		$table += '      <th>Group type</th>'
		$table += '      <th>Security enabled</th>'
		$table += '      <th>Membership type</th>'
		$table += '      <th>Membership rule</th>'
		$table += '    </tr>'
		$table += '  </thead>'
		$table += '  <tbody>'
		$table += ($rows -join [Environment]::NewLine)
		$table += '  </tbody>'
		$table += '</table>'
		$table += '</div>'
		return ($table -join [Environment]::NewLine)
	}

	function New-LatestUserGroupMembershipTableHtml {
		param(
			[array]$Groups,
			[bool]$HasLatestUser
		)

		if (-not $HasLatestUser) {
			return '<p>No latest logged-on user available.</p>'
		}
		if (-not $Groups -or $Groups.Count -eq 0) {
			return '<p>No latest user group memberships resolved.</p>'
		}

		$encode = {
			param($value)
			if ($null -eq $value) { return '' }
			return [System.Net.WebUtility]::HtmlEncode([string]$value)
		}

		$rows = foreach ($group in ($Groups | Sort-Object -Property displayName)) {
			$display = & $encode $group.displayName
			$devices = & $encode $group.YodamiittiCustomGroupMembersCountDevices
			$users = & $encode $group.YodamiittiCustomGroupMembersCountUsers
			$type = & $encode $group.YodamiittiCustomGroupType
			$security = & $encode $group.securityEnabled
			$membershipType = & $encode $group.YodamiittiCustomMembershipType
			$rule = & $encode $group.membershipRule
			$descriptionTooltip = if ([string]::IsNullOrWhiteSpace($group.description)) { '' } else { & $encode $group.description }
			$descriptionAttr = if ($descriptionTooltip) { " title='$descriptionTooltip'" } else { '' }
			$ruleTooltip = if ([string]::IsNullOrWhiteSpace($group.membershipRule)) { '' } else { & $encode $group.membershipRule }
			$ruleAttr = if ($ruleTooltip) { " title='$ruleTooltip'" } else { '' }

			$rowClass = ''
			if ($group.YodamiittiCustomGroupType -eq 'DirectoryRole') { $rowClass = 'role-directory' }
			if ($group.displayName -eq 'Global Administrator') { $rowClass = 'role-globaladmin' }
			$rowClassAttr = if ($rowClass) { " class='$rowClass'" } else { '' }

			"<tr$rowClassAttr><td class='col-display'$descriptionAttr><strong>$display</strong></td><td class='center'>$devices</td><td class='center'>$users</td><td class='col-groupType'>$type</td><td class='center'>$security</td><td class='col-groupType'>$membershipType</td><td class='col-rule'$ruleAttr>$rule</td></tr>"
		}

		$table = @()
		$table += '<div class="table-scroll-wrapper">'
		$table += '<table class="data-table sortable group-membership-table latest-user-group-table" data-table="Latest user group memberships">'
		$table += '  <thead>'
		$table += '    <tr>'
		$table += '      <th>Display name</th>'
		$table += '      <th>Devices</th>'
		$table += '      <th>Users</th>'
		$table += '      <th>Group type</th>'
		$table += '      <th>Security enabled</th>'
		$table += '      <th>Membership type</th>'
		$table += '      <th>Membership rule</th>'
		$table += '    </tr>'
		$table += '  </thead>'
		$table += '  <tbody>'
		$table += ($rows -join [Environment]::NewLine)
		$table += '  </tbody>'
		$table += '</table>'
		$table += '</div>'
		return ($table -join [Environment]::NewLine)
	}

	function New-DeviceGroupMembershipTableHtml {
		param([array]$Groups)

		if (-not $Groups -or $Groups.Count -eq 0) {
			return '<p>No device group memberships found.</p>'
		}

		$encode = {
			param($value)
			if ($null -eq $value) { return '' }
			return [System.Net.WebUtility]::HtmlEncode([string]$value)
		}

		$rows = foreach ($group in ($Groups | Sort-Object -Property displayName)) {
			$display = & $encode $group.displayName
			$devices = & $encode $group.YodamiittiCustomGroupMembersCountDevices
			$users = & $encode $group.YodamiittiCustomGroupMembersCountUsers
			$type = & $encode $group.YodamiittiCustomGroupType
			$security = & $encode $group.securityEnabled
			$membershipType = & $encode $group.YodamiittiCustomMembershipType
			$rule = & $encode $group.membershipRule
			$descriptionTooltip = if ([string]::IsNullOrWhiteSpace($group.description)) { '' } else { & $encode $group.description }
			$descriptionAttr = if ($descriptionTooltip) { " title='$descriptionTooltip'" } else { '' }
			$ruleTooltip = if ([string]::IsNullOrWhiteSpace($group.membershipRule)) { '' } else { & $encode $group.membershipRule }
			$ruleAttr = if ($ruleTooltip) { " title='$ruleTooltip'" } else { '' }
			$rowClass = ''
			if ($group.YodamiittiCustomGroupType -eq 'DirectoryRole') { $rowClass = 'role-directory' }
			if ($group.displayName -eq 'Global Administrator') { $rowClass = 'role-globaladmin' }
			$rowClassAttr = if ($rowClass) { " class='$rowClass'" } else { '' }

			"<tr$rowClassAttr><td class='col-display'$descriptionAttr><strong>$display</strong></td><td class='center'>$devices</td><td class='center'>$users</td><td class='col-groupType'>$type</td><td class='center'>$security</td><td class='col-groupType'>$membershipType</td><td class='col-rule'$ruleAttr>$rule</td></tr>"
		}

		$table = @()
		$table += '<div class="table-scroll-wrapper">'
		$table += '<table class="data-table sortable group-membership-table device-group-table" data-table="Device group memberships">'
		$table += '  <thead>'
		$table += '    <tr>'
		$table += '      <th>Display name</th>'
		$table += '      <th>Devices</th>'
		$table += '      <th>Users</th>'
		$table += '      <th>Group type</th>'
		$table += '      <th>Security enabled</th>'
		$table += '      <th>Membership type</th>'
		$table += '      <th>Membership rule</th>'
		$table += '    </tr>'
		$table += '  </thead>'
		$table += '  <tbody>'
		$table += ($rows -join [Environment]::NewLine)
		$table += '  </tbody>'
		$table += '</table>'
		$table += '</div>'
		return ($table -join [Environment]::NewLine)
	}

	function New-AppAssignmentsTableHtml {
		param([array]$Assignments)

		if (-not $Assignments -or $Assignments.Count -eq 0) {
			return '<p>No application assignments resolved.</p>'
		}

		$encode = {
			param($value)
			if ($null -eq $value) { return '' }
			return [System.Net.WebUtility]::HtmlEncode([string]$value)
		}

		$tooltipAttr = {
			param($value)
			if ([string]::IsNullOrWhiteSpace($value)) { return '' }
			return " title=`"$(& $encode $value)`""
		}

		$table = @()
		$table += '<div class="table-scroll-wrapper">'
		$table += '<table class="apps-table sortable" data-table="Application assignments">'
		$table += '  <thead>'
		$table += '    <tr>'
		$table += '      <th>Context</th>'
		$table += '      <th>Application type</th>'
		$table += '      <th>Display name</th>'
		$table += '      <th>Version</th>'
		$table += '      <th>Intent</th>'
		$table += '      <th>Include/Exclude</th>'
		$table += '      <th>Install state</th>'
		$table += '      <th>Group type</th>'
		$table += '      <th>Assignment group</th>'
		$table += '      <th>Group members</th>'
		$table += '      <th>Filter</th>'
		$table += '      <th>Filter mode</th>'
		$table += '    </tr>'
		$table += '  </thead>'
		$table += '  <tbody>'

		foreach ($app in $Assignments) {
			$context = & $encode $app.context
			$odata = & $encode $app.odatatype
			$display = & $encode $app.displayName
			$version = & $encode $app.version
			$intent = & $encode $app.assignmentIntent
			$includeExclude = & $encode $app.IncludeExclude
			$installState = & $encode $app.installState
			$groupType = & $encode $app.YodamiittiCustomMembershipType
			$assignmentGroup = & $encode $app.assignmentGroup
			$groupMembers = & $encode $app.YodamiittiCustomGroupMembers
			$filter = & $encode $app.filter
			$filterMode = & $encode $app.filterMode

			$contextClass = if ($app.context -in @('_unknown','_Device/User')) { ' warning-cell' } else { '' }
			$includeClass = if ($app.IncludeExclude -eq 'Excluded') { ' warning-cell' } else { '' }
			$installClass = switch ($app.installState) {
				'failed' { ' danger-cell' }
				'installed' { ' success-cell' }
				'notApplicable' { ' warning-cell' }
				Default { '' }
			}
			$assignmentGroupClass = ''
			if ($app.assignmentGroup -eq 'Application does not have any assignments!') {
				$assignmentGroupClass = ' danger-cell'
			} elseif (($app.installState -eq 'notApplicable') -or ($app.IncludeExclude -eq 'Excluded') -or ($app.assignmentGroup -eq 'unknown (possible nested group or removed assignment)')) {
				$assignmentGroupClass = ' warning-cell'
			}
			$filterHighlight = if ($app.installState -eq 'notApplicable') { ' warning-cell' } else { '' }

			$contextTooltipAttr = & $tooltipAttr $app.contextToolTip
			
			# Build displayNameToolTip - always include description, optionally add Win32LobApp details
			$enhancedDisplayTooltip = $app.displayNameToolTip
			
			# Add Win32LobApp or Win32CatalogApp details if available
			if ($app.id -and $script:GUIDHashtable.ContainsKey($app.id)) {
				$appData = $script:GUIDHashtable[$app.id]
				$odataType = $appData.Object.'@odata.type'
				if ($odataType -eq '#microsoft.graph.win32LobApp' -or $odataType -eq '#microsoft.graph.win32CatalogApp') {
					$win32Details = ConvertTo-ReadableWin32LobApp -AppData $appData.Object -ExtendedReport:$script:useExtendedReport
					if ($win32Details) {
						if ([string]::IsNullOrWhiteSpace($enhancedDisplayTooltip)) {
							$enhancedDisplayTooltip = $win32Details
						} else {
							$enhancedDisplayTooltip = "$enhancedDisplayTooltip`n`n$win32Details"
						}
					}
			}
			elseif ($odataType -eq '#microsoft.graph.macOSDmgApp') {
				$macOSDetails = ConvertTo-ReadableMacOSDmgApp -AppData $appData.Object
				if ($macOSDetails) {
					if ([string]::IsNullOrWhiteSpace($enhancedDisplayTooltip)) {
						$enhancedDisplayTooltip = $macOSDetails
					} else {
						$enhancedDisplayTooltip = "$enhancedDisplayTooltip`n`n$macOSDetails"
					}
				}
			}
			elseif ($odataType -eq '#microsoft.graph.macOSPkgApp') {
				$macOSPkgDetails = ConvertTo-ReadableMacOSPkgApp -AppData $appData.Object
				if ($macOSPkgDetails) {
					if ([string]::IsNullOrWhiteSpace($enhancedDisplayTooltip)) {
						$enhancedDisplayTooltip = $macOSPkgDetails
					} else {
						$enhancedDisplayTooltip = "$enhancedDisplayTooltip`n`n$macOSPkgDetails"
					}
				}
			}
			elseif ($odataType -eq '#microsoft.graph.iosVppApp') {
				$iosVppDetails = ConvertTo-ReadableIosVppApp -AppData $appData.Object
				if ($iosVppDetails) {
					if ([string]::IsNullOrWhiteSpace($enhancedDisplayTooltip)) {
						$enhancedDisplayTooltip = $iosVppDetails
					} else {
						$enhancedDisplayTooltip = "$enhancedDisplayTooltip`n`n$iosVppDetails"
					}
				}
			}
			elseif ($odataType -eq '#microsoft.graph.macOsVppApp') {
				$macOsVppDetails = ConvertTo-ReadableMacOsVppApp -AppData $appData.Object
				if ($macOsVppDetails) {
					if ([string]::IsNullOrWhiteSpace($enhancedDisplayTooltip)) {
						$enhancedDisplayTooltip = $macOsVppDetails
					} else {
						$enhancedDisplayTooltip = "$enhancedDisplayTooltip`n`n$macOsVppDetails"
					}
				}
			}
			elseif ($odataType -eq '#microsoft.graph.webApp') {
				$webAppDetails = ConvertTo-ReadableWebApp -AppData $appData.Object
				if ($webAppDetails) {
					if ([string]::IsNullOrWhiteSpace($enhancedDisplayTooltip)) {
						$enhancedDisplayTooltip = $webAppDetails
					} else {
						$enhancedDisplayTooltip = "$enhancedDisplayTooltip`n`n$webAppDetails"
					}
				}
			}
			elseif ($odataType -eq '#microsoft.graph.winGetApp') {
				$winGetDetails = ConvertTo-ReadableWinGetApp -AppData $appData.Object
				if ($winGetDetails) {
					if ([string]::IsNullOrWhiteSpace($enhancedDisplayTooltip)) {
						$enhancedDisplayTooltip = $winGetDetails
					} else {
						$enhancedDisplayTooltip = "$enhancedDisplayTooltip`n`n$winGetDetails"
					}
				}
			}
		}
		
		$displayTooltipAttr = & $tooltipAttr $enhancedDisplayTooltip
		
		$assignmentTooltipAttr = & $tooltipAttr $app.AssignmentGroupToolTip
		$filterTooltipAttr = & $tooltipAttr $app.filterTooltip

		$rowCells = @()
			$rowCells += "<td class='col-context center$contextClass'$contextTooltipAttr>$context</td>"
			$rowCells += "<td class='col-odatatype'>$odata</td>"
			$rowCells += "<td class='col-display'$displayTooltipAttr><strong>$display</strong></td>"
			$rowCells += "<td class='shrink'>$version</td>"
			$rowCells += "<td class='shrink center'>$intent</td>"
			$rowCells += "<td class='shrink center$includeClass'>$includeExclude</td>"
			$rowCells += "<td class='shrink center$installClass'>$installState</td>"
			$rowCells += "<td class='shrink'>$groupType</td>"
			$rowCells += "<td class='col-group$assignmentGroupClass'$assignmentTooltipAttr>$assignmentGroup</td>"
			$rowCells += "<td class='col-members right'>$groupMembers</td>"
			$rowCells += "<td class='col-filter$filterHighlight'$filterTooltipAttr>$filter</td>"
			$rowCells += "<td class='col-filterMode$filterHighlight'>$filterMode</td>"

			$table += "    <tr>$([string]::Join('', $rowCells))</tr>"
		}

		$table += '  </tbody>'
		$table += '</table>'
		$table += '</div>'
		return ($table -join [Environment]::NewLine)
	}

	function New-ConfigAssignmentsTableHtml {
		param([array]$Policies)

		if (-not $Policies -or $Policies.Count -eq 0) {
			return '<p>No configuration policy data returned.</p>'
		}

		$encode = {
			param($value)
			if ($null -eq $value) { return '' }
			return [System.Net.WebUtility]::HtmlEncode([string]$value)
		}

		$tooltipAttr = {
			param($value)
			if ([string]::IsNullOrWhiteSpace($value)) { return '' }
			return " title=`"$(& $encode $value)`""
		}

		$table = @()
		$table += '<div class="table-scroll-wrapper">'
		$table += '<table class="config-table sortable" data-table="Configuration policies">'
		$table += '  <thead>'
		$table += '    <tr>'
		$table += '      <th>Context</th>'
		$table += '      <th>Configuration type</th>'
		$table += '      <th>Display name</th>'
		$table += '      <th>User principal name</th>'
		$table += '      <th>Include/Exclude</th>'
		$table += '      <th>State</th>'
		$table += '      <th>Group type</th>'
		$table += '      <th>Assignment group</th>'
		$table += '      <th>Group members</th>'
		$table += '      <th>Filter</th>'
		$table += '      <th>Filter mode</th>'
		$table += '    </tr>'
		$table += '  </thead>'
		$table += '  <tbody>'

		foreach ($policy in $Policies) {
			$context = & $encode $policy.context
			$odata = & $encode $policy.odatatype
			$display = & $encode $policy.displayName
			$upn = & $encode $policy.userPrincipalName
			$includeExclude = & $encode $policy.IncludeExclude
			$state = & $encode $policy.state
			$groupType = & $encode $policy.YodamiittiCustomMembershipType
			$assignmentGroup = & $encode $policy.assignmentGroup
			$groupMembers = & $encode $policy.YodamiittiCustomGroupMembers
			$filter = & $encode $policy.filter
			$filterMode = & $encode $policy.filterMode

			$contextClass = if ($policy.context -in @('_unknown','_Device/User')) { ' warning-cell' } else { '' }
			$includeClass = if ($policy.IncludeExclude -eq 'Excluded') { ' warning-cell' } else { '' }
			$stateClass = switch ($policy.state) {
				'Succeeded' { ' success-cell' }
				'Conflict' { ' danger-cell' }
				'Error' { ' danger-cell' }
				'Not applicable' { ' warning-cell' }
				Default { '' }
			}
			$assignmentGroupClass = ''
			if ($policy.assignmentGroup -eq 'Policy does not have any assignments!') {
				$assignmentGroupClass = ' danger-cell'
			} elseif (($policy.state -eq 'Not applicable') -or ($policy.IncludeExclude -eq 'Excluded') -or ($policy.assignmentGroup -eq 'unknown (possible user targeted group, nested group or removed assignment)')) {
				$assignmentGroupClass = ' warning-cell'
			}
			$filterHighlight = if ($policy.state -eq 'Not applicable') { ' warning-cell' } else { '' }

			$contextTooltipAttr = & $tooltipAttr $policy.contextToolTip
			
			# Build displayNameToolTip - always include description, optionally add Settings Catalog details or OMA settings
			$enhancedDisplayTooltip = $policy.displayNameToolTip
			
			# Add Settings Catalog details if available (ExtendedReport)
			if ($policy.id -and $script:GUIDHashtable.ContainsKey($policy.id)) {
				$policyData = $script:GUIDHashtable[$policy.id]
				if ($policyData.settingsCatalogDetails) {
					# Combine description with Settings Catalog details
					if ([string]::IsNullOrWhiteSpace($enhancedDisplayTooltip)) {
						$enhancedDisplayTooltip = "--- Configuration settings ---`n`n$($policyData.settingsCatalogDetails)"
					} else {
						$enhancedDisplayTooltip = "$enhancedDisplayTooltip`n`n--- Configuration settings ---`n`n$($policyData.settingsCatalogDetails)"
					}
				}
				
			# Add OMA settings if available (windows10CustomConfiguration policies)
			if ($policyData.Object.'@odata.type' -eq '#microsoft.graph.windows10CustomConfiguration' -and $policyData.Object.omaSettings) {
				$omaDetails = ConvertTo-ReadableOmaSettings -OmaSettings $policyData.Object.omaSettings
				if ($omaDetails) {
					if ([string]::IsNullOrWhiteSpace($enhancedDisplayTooltip)) {
						$enhancedDisplayTooltip = "--- OMA-URI Settings ---`n`n$omaDetails"
					} else {
						$enhancedDisplayTooltip = "$enhancedDisplayTooltip`n`n--- OMA-URI Settings ---`n`n$omaDetails"
					}
				}
			}
			
			# Add macOS Custom Configuration details if available
			if ($policyData.Object.'@odata.type' -eq '#microsoft.graph.macOSCustomConfiguration') {
				$macOSCustomDetails = ConvertTo-ReadableMacOSCustomConfiguration -PolicyData $policyData.Object
				if ($macOSCustomDetails) {
					if ([string]::IsNullOrWhiteSpace($enhancedDisplayTooltip)) {
						$enhancedDisplayTooltip = $macOSCustomDetails
					} else {
						$enhancedDisplayTooltip = "$enhancedDisplayTooltip`n`n$macOSCustomDetails"
					}
				}
			}
			
			# Add macOS Custom App Configuration (plist) details if available
			if ($policyData.Object.'@odata.type' -eq '#microsoft.graph.macOSCustomAppConfiguration') {
				$macOSAppConfigDetails = ConvertTo-ReadableMacOSCustomAppConfiguration -PolicyData $policyData.Object
				if ($macOSAppConfigDetails) {
					if ([string]::IsNullOrWhiteSpace($enhancedDisplayTooltip)) {
						$enhancedDisplayTooltip = $macOSAppConfigDetails
					} else {
						$enhancedDisplayTooltip = "$enhancedDisplayTooltip`n`n$macOSAppConfigDetails"
					}
				}
			}
		}
		
		$displayTooltipAttr = & $tooltipAttr $enhancedDisplayTooltip
		
		$assignmentTooltipAttr = & $tooltipAttr $policy.AssignmentGroupToolTip
		$filterTooltipAttr = & $tooltipAttr $policy.filterTooltip

		$rowCells = @()
			$rowCells += "<td class='col-context center$contextClass'$contextTooltipAttr>$context</td>"
			$rowCells += "<td class='col-odatatype'>$odata</td>"
			$rowCells += "<td class='col-display'$displayTooltipAttr><strong>$display</strong></td>"
			$rowCells += "<td class='col-upn'>$upn</td>"
			$rowCells += "<td class='shrink center$includeClass'>$includeExclude</td>"
			$rowCells += "<td class='col-state center$stateClass'>$state</td>"
			$rowCells += "<td class='col-groupType center'>$groupType</td>"
			$rowCells += "<td class='col-group$assignmentGroupClass'$assignmentTooltipAttr>$assignmentGroup</td>"
			$rowCells += "<td class='col-members right'>$groupMembers</td>"
			$rowCells += "<td class='col-filter$filterHighlight'$filterTooltipAttr>$filter</td>"
			$rowCells += "<td class='col-filterMode$filterHighlight'>$filterMode</td>"

			$table += "    <tr>$([string]::Join('', $rowCells))</tr>"
		}

		$table += '  </tbody>'
		$table += '</table>'
		$table += '</div>'
		return ($table -join [Environment]::NewLine)
	}

	function New-RemediationScriptsTableHtml {
		param([array]$Scripts)

		if (-not $Scripts -or $Scripts.Count -eq 0) {
			return '<p>No remediation scripts found for this device.</p>'
		}

		$encode = {
			param($value)
			if ($null -eq $value) { return '' }
			return [System.Net.WebUtility]::HtmlEncode([string]$value)
		}

		$tooltipAttr = {
			param($value)
			if ([string]::IsNullOrWhiteSpace($value)) { return '' }
			return " title=`"$(& $encode $value)`""
		}

		$table = @()
		$table += '<div class="table-scroll-wrapper">'
		$table += '<table class="config-table sortable" data-table="Remediation scripts">'
		$table += '  <thead>'
		$table += '    <tr>'
		$table += '      <th>Context</th>'
		$table += '      <th>Script type</th>'
		$table += '      <th>Script name</th>'
		$table += '      <th>Detection status</th>'
		$table += '      <th>Remediation status</th>'
		$table += '      <th>User principal name</th>'
		$table += '      <th>Status updated</th>'
		$table += '      <th>Group type</th>'
		$table += '      <th>Assignment group</th>'
		$table += '      <th>Group members</th>'
		$table += '      <th>Filter</th>'
		$table += '      <th>Filter mode</th>'
		$table += '    </tr>'
		$table += '  </thead>'
		$table += '  <tbody>'

		foreach ($script in $Scripts) {
			$context = & $encode $script.context
			$scriptType = & $encode $script.scriptType
			$displayName = & $encode $script.displayName
			$detectionStatus = & $encode $script.detectionStatus
			$remediationStatus = & $encode $script.remediationStatus
			$upn = & $encode $script.userPrincipalName
			$statusUpdateTime = & $encode $script.statusUpdateTime
			$groupType = & $encode $script.groupType
			$assignmentGroup = & $encode $script.assignmentGroup
			$groupMembers = & $encode $script.groupMembers
			$filter = & $encode $script.filter
			$filterMode = & $encode $script.filterMode

			$contextClass = if ($script.context -in @('_unknown','_Device/User')) { ' warning-cell' } else { '' }
			$detectionClass = switch ($script.detectionStatus) {
				'Without issues' { ' success-cell' }
				'With issues' { ' warning-cell' }
				'Not applicable' { ' neutral-cell' }
				Default { '' }
			}
			$remediationClass = switch ($script.remediationStatus) {
				'Issue fixed' { ' success-cell' }
				'With issues' { ' danger-cell' }
				'Not run' { ' neutral-cell' }
				Default { '' }
			}
			$assignmentGroupClass = if ($script.assignmentGroup -eq 'No assignments') { ' danger-cell' } else { '' }
			$filterHighlight = if ($script.filter) { ' warning-cell' } else { '' }

			$detectionTooltipAttr = & $tooltipAttr $script.detectionStatusTooltip
			$remediationTooltipAttr = & $tooltipAttr $script.remediationStatusTooltip
			$timeTooltipAttr = & $tooltipAttr $script.statusUpdateTimeTooltip
			$assignmentTooltipAttr = & $tooltipAttr $script.assignmentGroupTooltip
			$filterTooltipAttr = & $tooltipAttr $script.filterTooltip

		# Get script content tooltip for scripts if available
		$scriptNameTooltipAttr = ''
		if ($script.id) {
			if ($script.scriptType -like 'Platform (Windows)') {
				# Platform script - single script content
				if ($script:GUIDHashtable.ContainsKey($script.id)) {
					$scriptObject = $script:GUIDHashtable[$script.id]
					if ($scriptObject.scriptContentClearText) {
						$scriptNameTooltipAttr = & $tooltipAttr $scriptObject.scriptContentClearText
					}
				}
			}
			elseif ($script.scriptType -like 'Platform (macOS)') {
				# macOS shell script - show script content if available, otherwise show metadata
				if ($script:GUIDHashtable.ContainsKey($script.id)) {
					$scriptObject = $script:GUIDHashtable[$script.id]
					
					# Prefer script content if downloaded (ExtendedReport mode)
					if ($scriptObject.scriptContentClearText) {
						$scriptNameTooltipAttr = & $tooltipAttr $scriptObject.scriptContentClearText
					}
					else {
						# Fall back to metadata
						$tooltipParts = @()
						
						if ($scriptObject.fileName) {
							$tooltipParts += "File name: $($scriptObject.fileName)"
						}
						if ($scriptObject.description) {
							$tooltipParts += "Description: $($scriptObject.description)"
						}
						if ($scriptObject.runAsAccount) {
							$tooltipParts += "Run as: $($scriptObject.runAsAccount)"
						}
						if ($scriptObject.executionFrequency) {
							# Convert ISO 8601 duration to readable format
							$execFreq = $scriptObject.executionFrequency
							if ($execFreq -match 'PT(\d+)M') {
								$execFreq = "$($Matches[1]) minutes"
							}
							elseif ($execFreq -match 'PT(\d+)H') {
								$execFreq = "$($Matches[1]) hours"
							}
							elseif ($execFreq -match 'P(\d+)D') {
								$execFreq = "$($Matches[1]) days"
							}
							$tooltipParts += "Execution frequency: $execFreq"
						}
						if ($scriptObject.retryCount) {
							$tooltipParts += "Retry count: $($scriptObject.retryCount)"
						}
						if ($null -ne $scriptObject.blockExecutionNotifications) {
							$blockNotifications = if ($scriptObject.blockExecutionNotifications) { 'Yes' } else { 'No' }
							$tooltipParts += "Block notifications: $blockNotifications"
						}
						
						if ($tooltipParts.Count -gt 0) {
							$combinedTooltip = $tooltipParts -join "`n"
							$scriptNameTooltipAttr = & $tooltipAttr $combinedTooltip
						}
					}
				}
			}
			elseif ($script.scriptType -like 'Platform (Linux)') {
				# Linux bash script - show metadata
				if ($script:GUIDHashtable.ContainsKey($script.id)) {
					$scriptObject = $script:GUIDHashtable[$script.id]
					$tooltipParts = @()
					
					if ($scriptObject.name) {
						$tooltipParts += "Name: $($scriptObject.name)"
					}
					if ($scriptObject.description) {
						$tooltipParts += "Description: $($scriptObject.description)"
					}
					if ($scriptObject.platforms) {
						$tooltipParts += "Platform: $($scriptObject.platforms)"
					}
					if ($scriptObject.technologies) {
						$tooltipParts += "Technology: $($scriptObject.technologies)"
					}
					if ($scriptObject.settingCount) {
						$tooltipParts += "Setting count: $($scriptObject.settingCount)"
					}
					
					if ($tooltipParts.Count -gt 0) {
						$combinedTooltip = $tooltipParts -join "`n"
						$scriptNameTooltipAttr = & $tooltipAttr $combinedTooltip
					}
				}
			}
			elseif ($script.scriptType -eq 'Remediation') {
				# Remediation script - can have detection and/or remediation scripts
				if ($script:GUIDHashtable.ContainsKey($script.id)) {
					$scriptObject = $script:GUIDHashtable[$script.id]
					$tooltipParts = @()
					
					if ($scriptObject.detectionScriptContentClearText) {
						$tooltipParts += "Detection script:`n`n$($scriptObject.detectionScriptContentClearText)"
					}
					
					if ($scriptObject.remediateScriptContentClearText) {
						if ($tooltipParts.Count -gt 0) {
							$tooltipParts += "`n`n---`n`n"
						}
						$tooltipParts += "Remediation script:`n`n$($scriptObject.remediateScriptContentClearText)"
					}
					
					if ($tooltipParts.Count -gt 0) {
						$combinedTooltip = $tooltipParts -join ''
						$scriptNameTooltipAttr = & $tooltipAttr $combinedTooltip
					}
				}
			}
		}
			$rowCells = @()
			$rowCells += "<td class='col-context center$contextClass'>$context</td>"
			$rowCells += "<td class='col-scriptType'>$scriptType</td>"
			$rowCells += "<td class='col-display'$scriptNameTooltipAttr><strong>$displayName</strong></td>"
			$rowCells += "<td class='col-state center$detectionClass'$detectionTooltipAttr>$detectionStatus</td>"
			$rowCells += "<td class='col-state center$remediationClass'$remediationTooltipAttr>$remediationStatus</td>"
			$rowCells += "<td class='col-upn'>$upn</td>"
			$rowCells += "<td class='col-time'$timeTooltipAttr>$statusUpdateTime</td>"
			$rowCells += "<td class='col-groupType center'>$groupType</td>"
			$rowCells += "<td class='col-group$assignmentGroupClass'$assignmentTooltipAttr>$assignmentGroup</td>"
			$rowCells += "<td class='col-members right'>$groupMembers</td>"
			$rowCells += "<td class='col-filter$filterHighlight'$filterTooltipAttr>$filter</td>"
			$rowCells += "<td class='col-filterMode$filterHighlight'>$filterMode</td>"

			$table += "    <tr>$([string]::Join('', $rowCells))</tr>"
		}

		$table += '  </tbody>'
		$table += '</table>'
		$table += '</div>'
		return ($table -join [Environment]::NewLine)
	}

	$appTable = New-AppAssignmentsTableHtml -Assignments $appAssignments
	$configTable = New-ConfigAssignmentsTableHtml -Policies $configPolicies
	$remediationTable = New-RemediationScriptsTableHtml -Scripts $remediationScripts
	$primaryGroupTable = New-PrimaryUserGroupMembershipTableHtml -Groups $primaryUserGroups -HasPrimaryUser:([bool]$primaryUser)
	$latestGroupTable = New-LatestUserGroupMembershipTableHtml -Groups $latestUserGroups -HasLatestUser:([bool]$latestUser)

	$groupTable = New-DeviceGroupMembershipTableHtml -Groups $deviceGroups

	$autopilotJson = if ($autopilotDetail) {
		$autopilotDetail | Select-Object -Property * -ExcludeProperty DeploymentProfileDetail | ConvertTo-Json -Depth 4
	} else { $null }
	$autopilotDeploymentProfileJson = if ($autopilotDetail) {
		if ($autopilotDetail.DeploymentProfileDetail) {
			$autopilotDetail.DeploymentProfileDetail | ConvertTo-Json -Depth 6
		}
		elseif ($autopilotDetail.deploymentProfile) {
			$autopilotDetail.deploymentProfile | ConvertTo-Json -Depth 6
		}
		else { $null }
	} else { $null }
	$espJson = if ($espContext -and $espContext.Detail) { $espContext.Detail | ConvertTo-Json -Depth 6 } else { $null }
	$autopilotDevicePrepJson = if ($script:AutopilotDevicePreparationPolicyWithAssignments) { $script:AutopilotDevicePreparationPolicyWithAssignments | ConvertTo-Json -Depth 6 } else { $null }
	$appleEnrollmentProfileJson = if ($script:AppleEnrollmentProfileDetails) { $script:AppleEnrollmentProfileDetails | ConvertTo-Json -Depth 6 } else { $null }

	$deviceJson = $device | ConvertTo-Json -Depth 6
	$entraDeviceJson = if ($azureDevice) {
		$sanitized = $azureDevice.PSObject.Copy()
		if ($sanitized.alternativeSecurityIds) {
			foreach ($secId in $sanitized.alternativeSecurityIds) {
				if ($secId.key) { $secId.key = "***CENSORED***" }
			}
		}
		$sanitized | ConvertTo-Json -Depth 6
	} else { $null }
	$primaryUserJson = if ($primaryUser) {
		$sanitized = $primaryUser.PSObject.Copy()
		if ($sanitized.deviceKeys) {
			foreach ($key in $sanitized.deviceKeys) {
				if ($key.keyMaterial) { $key.keyMaterial = "***CENSORED***" }
			}
		}
		$sanitized | ConvertTo-Json -Depth 6
	} else { $null }
	$latestUserJson = if ($latestUser) {
		$sanitized = $latestUser.PSObject.Copy()
		if ($sanitized.deviceKeys) {
			foreach ($key in $sanitized.deviceKeys) {
				if ($key.keyMaterial) { $key.keyMaterial = "***CENSORED***" }
			}
		}
		$sanitized | ConvertTo-Json -Depth 6
	} else { $null }

	# Get current time and domain
	$reportRunTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
	$userDomain = if ($script:ConnectedUser -and $script:ConnectedUser -match '@(.+)$') { $Matches[1] } else { '' }
	$tenantInfo = if ($script:TenantDisplayName) { $script:TenantDisplayName } else { $script:TenantId }
	
	# Determine report type based on parameters
	if ($SkipAssignments) {
		$reportType = 'Minimal Report'
	} elseif ($ExtendedReport) {
		$reportType = 'Extended Report'
	} else {
		$reportType = 'Normal Report'
	}

	$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Intune Device Report - $($device.deviceName)</title>
<style>$css</style>
</head>
<body>
<div class="report-header">
  <div class="left">Report run: $reportRunTime | Tenant: $tenantInfo | $reportType</div>
  <div class="center">IntuneDeviceDetailsHTML Version $Version</div>
  <div class="right"><a href="https://github.com/petripaavola/IntuneDeviceDetailsGUI" target="_blank" style="color: #0066cc; text-decoration: none; font-weight: 700;">Download from GitHub</a></div>
</div>
<div class="page">
  <div class="card">
    <div class="device-info-layout">
      <div class="badge-column">
        $complianceBadge
        $autopilotBadge
        $encryptionBadge
      </div>
      <div>
        $deviceInfoCards
      </div>
    </div>
  </div>

	<div class="card">
		<h3>Quick table search</h3>
		<div class="table-search">
			<label for="table-search">Filter rows</label>
			<input id="table-search" type="text" placeholder="Type to filter (use commas for multiple terms)" />
			<button class="clear-btn" onclick="document.getElementById('table-search').value='';document.getElementById('table-search').dispatchEvent(new Event('input'));">✕</button>
		</div>
		<p style="margin-top:8px; margin-bottom:0; font-size:11px; color:#64748b;">💡 Click any table row to view detailed information in a popup window</p>
	</div>

  <div class="card">
    <div class="card-header">
      <h3>Application assignments</h3>
      <div class="card-controls">
        <button class="card-control-btn" data-action="minimize" title="Minimize">−</button>
        <button class="card-control-btn active" data-action="normal" title="Normal size">□</button>
        <button class="card-control-btn" data-action="fullsize" title="Full size">⬜</button>
      </div>
    </div>
    <div class="card-body">
      $appTable
      $(if ($Context.AppAssignments.UnknownAssignments) { '<p><em>Some assignments could not be resolved (possible nested groups or stale cache).</em></p>' } else { '' })
    </div>
  </div>

  <div class="card">
    <div class="card-header">
      <h3>Configuration policies</h3>
      <div class="card-controls">
        <button class="card-control-btn" data-action="minimize" title="Minimize">−</button>
        <button class="card-control-btn active" data-action="normal" title="Normal size">□</button>
        <button class="card-control-btn" data-action="fullsize" title="Full size">⬜</button>
      </div>
    </div>
    <div class="card-body">
      $configTable
    </div>
  </div>

  <div class="card">
    <div class="card-header">
      <h3>Remediation and platform scripts</h3>
      <div class="card-controls">
        <button class="card-control-btn" data-action="minimize" title="Minimize">−</button>
        <button class="card-control-btn active" data-action="normal" title="Normal size">□</button>
        <button class="card-control-btn" data-action="fullsize" title="Full size">⬜</button>
      </div>
    </div>
    <div class="card-body">
      $remediationTable
    </div>
  </div>

  <div class="card">
    <div class="card-header">
      <h3>Group Memberships</h3>
      <div class="card-controls">
        <button class="card-control-btn" data-action="minimize" title="Minimize">−</button>
        <button class="card-control-btn active" data-action="normal" title="Normal size">□</button>
        <button class="card-control-btn" data-action="fullsize" title="Full size">⬜</button>
      </div>
    </div>
    <div class="card-body">
      <div class="tabs">
        <button class="tab-button active" data-tab="device-groups">Device Groups</button>
        <button class="tab-button" data-tab="primary-user-groups">Primary User Groups</button>
        <button class="tab-button" data-tab="latest-user-groups">Latest Logon User Groups</button>
      </div>
      <div class="tab-content active" id="device-groups-tab">
        $groupTable
      </div>
      <div class="tab-content" id="primary-user-groups-tab">
        $primaryGroupTable
      </div>
      <div class="tab-content" id="latest-user-groups-tab">
        $latestGroupTable
      </div>
    </div>
  </div>

  <div class="card">
    <h3>Raw JSON Data</h3>
    <div class="tabs">
      <button class="tab-button active" data-tab="device">Intune Device</button>
      $(if ($entraDeviceJson) { '<button class="tab-button" data-tab="entra-device">Entra Device</button>' } else { '' })
      $(if ($autopilotJson) { '<button class="tab-button" data-tab="autopilot">Autopilot Device</button>' } else { '' })
      $(if ($primaryUserJson) { '<button class="tab-button" data-tab="user">Primary User</button>' } else { '' })
      $(if ($latestUserJson) { '<button class="tab-button" data-tab="latest-user">Latest Logon User</button>' } else { '' })
      $(if ($autopilotDeploymentProfileJson) { '<button class="tab-button" data-tab="autopilot-profile">Autopilot Deployment Profile</button>' } else { '' })
      $(if ($autopilotDevicePrepJson) { '<button class="tab-button" data-tab="autopilot-device-prep">Autopilot Device Preparation</button>' } else { '' })
      $(if ($espJson) { '<button class="tab-button" data-tab="esp">Enrollment Status Page</button>' } else { '' })
      $(if ($appleEnrollmentProfileJson) { '<button class="tab-button" data-tab="apple-enrollment">Apple Enrollment Profile</button>' } else { '' })
    </div>
    <details>
      <summary>Expand</summary>
      <div class="tab-content active" id="device-tab">
        <pre>$deviceJson</pre>
      </div>
      $(if ($entraDeviceJson) { "<div class=`"tab-content`" id=`"entra-device-tab`"><pre>$entraDeviceJson</pre></div>" } else { '' })
      $(if ($autopilotJson) { "<div class=`"tab-content`" id=`"autopilot-tab`"><pre>$autopilotJson</pre></div>" } else { '' })
      $(if ($primaryUserJson) { "<div class=`"tab-content`" id=`"user-tab`"><pre>$primaryUserJson</pre></div>" } else { '' })
      $(if ($latestUserJson) { "<div class=`"tab-content`" id=`"latest-user-tab`"><pre>$latestUserJson</pre></div>" } else { '' })
      $(if ($autopilotDeploymentProfileJson) { "<div class=`"tab-content`" id=`"autopilot-profile-tab`"><pre>$autopilotDeploymentProfileJson</pre></div>" } else { '' })
      $(if ($autopilotDevicePrepJson) { "<div class=`"tab-content`" id=`"autopilot-device-prep-tab`"><pre>$autopilotDevicePrepJson</pre></div>" } else { '' })
      $(if ($espJson) { "<div class=`"tab-content`" id=`"esp-tab`"><pre>$espJson</pre></div>" } else { '' })
      $(if ($appleEnrollmentProfileJson) { "<div class=`"tab-content`" id=`"apple-enrollment-tab`"><pre>$appleEnrollmentProfileJson</pre></div>" } else { '' })
    </details>
  </div>
</div>

<div class="row-details-modal" id="rowDetailsModal">
  <div class="row-details-content">
    <div class="row-details-header">
      <h2>Row Details</h2>
      <button class="row-details-close" onclick="closeRowDetails()">Close</button>
    </div>
    <div class="row-details-body" id="rowDetailsBody">
    </div>
  </div>
</div>

<script>
(function(){
	const getCellValue = (row, index) => {
		const cell = row.cells[index];
		if (!cell) { return ''; }
		return cell.textContent.trim();
	};

	const parseValue = (value) => {
		const number = parseFloat(value.replace(/,/g,''));
		return isNaN(number) ? value.toLowerCase() : number;
	};

	const sortTable = (table, columnIndex, ascending) => {
		const tbody = table.tBodies[0];
		if (!tbody) { return; }
		const rows = Array.from(tbody.querySelectorAll('tr'));
		rows.sort((a, b) => {
			const aVal = parseValue(getCellValue(a, columnIndex));
			const bVal = parseValue(getCellValue(b, columnIndex));
			if (aVal < bVal) { return ascending ? -1 : 1; }
			if (aVal > bVal) { return ascending ? 1 : -1; }
			return 0;
		});
		rows.forEach(row => tbody.appendChild(row));
	};

	const attachSorting = (table) => {
		const headers = table.querySelectorAll('th');
		headers.forEach((header, index) => {
			header.addEventListener('click', () => {
				const current = header.getAttribute('data-sort-dir');
				const nextDir = current === 'asc' ? 'desc' : 'asc';
				headers.forEach(h => h.removeAttribute('data-sort-dir'));
				header.setAttribute('data-sort-dir', nextDir);
				sortTable(table, index, nextDir === 'asc');
			});
		});
	};

	const filterTables = (query) => {
		const rows = document.querySelectorAll('table.sortable tbody tr');
		const normalized = query.trim().toLowerCase();
		
		rows.forEach(row => {
			if (!normalized) {
				row.style.display = '';
				return;
			}
			
			// Search in both text content and title attributes (tooltips)
			const text = row.textContent.toLowerCase();
			const tooltips = Array.from(row.querySelectorAll('[title]'))
				.map(el => el.getAttribute('title').toLowerCase())
				.join(' ');
			const searchableContent = text + ' ' + tooltips;
			
			// Split by comma and trim each search term
			const searchTerms = normalized.split(',').map(term => term.trim()).filter(term => term.length > 0);
			
			// If no valid search terms after filtering, show all
			if (searchTerms.length === 0) {
				row.style.display = '';
				return;
			}
			
			// Any search term can match (OR logic)
			const anyTermMatches = searchTerms.some(term => searchableContent.includes(term));
			row.style.display = anyTermMatches ? '' : 'none';
		});
	};

	document.addEventListener('DOMContentLoaded', () => {
		document.querySelectorAll('table.sortable').forEach(attachSorting);
		const searchInput = document.getElementById('table-search');
		if (searchInput) {
			searchInput.addEventListener('input', () => filterTables(searchInput.value));
		}

		// Tab switching
		document.querySelectorAll('.tab-button').forEach(button => {
			button.addEventListener('click', () => {
				const tabName = button.getAttribute('data-tab');
				const parentCard = button.closest('.card');
				parentCard.querySelectorAll('.tab-button').forEach(b => b.classList.remove('active'));
				parentCard.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
				button.classList.add('active');
				document.getElementById(tabName + '-tab').classList.add('active');
			});
		});

		// Card size controls
		document.querySelectorAll('.card-control-btn').forEach(btn => {
			btn.addEventListener('click', (e) => {
				const action = btn.getAttribute('data-action');
				const card = btn.closest('.card');
				const allBtns = card.querySelectorAll('.card-control-btn');
				
				allBtns.forEach(b => b.classList.remove('active'));
				btn.classList.add('active');
				
				card.classList.remove('minimized', 'fullsize');
				if (action === 'minimize') {
					card.classList.add('minimized');
				} else if (action === 'fullsize') {
					card.classList.add('fullsize');
					// Scroll the card to the top of the viewport
					setTimeout(() => {
						card.scrollIntoView({ behavior: 'smooth', block: 'start' });
					}, 50);
				}
			});
		});

		// Tooltip with selectable text
		const tooltipBox = document.createElement('div');
		tooltipBox.className = 'tooltip-box';
		document.body.appendChild(tooltipBox);

		let currentCard = null;

		document.querySelectorAll('.info-card[data-tooltip]').forEach(card => {
			card.addEventListener('mouseenter', (e) => {
				currentCard = card;
				const rect = card.getBoundingClientRect();
				tooltipBox.textContent = card.getAttribute('data-tooltip');
				tooltipBox.style.display = 'block';
				
				// Position and measure
				setTimeout(() => {
					const tooltipRect = tooltipBox.getBoundingClientRect();
					const viewportWidth = window.innerWidth;
					const viewportHeight = window.innerHeight;
					let leftPos = rect.left;
					let topPos = rect.bottom + 8;
					
					// Check if tooltip would overflow bottom of viewport
					if (topPos + tooltipRect.height > viewportHeight - 10) {
						// Position tooltip ABOVE the card instead
						topPos = rect.top - tooltipRect.height - 8;
						
						// If it would overflow the top as well, position at top of viewport
						if (topPos < 10) {
							topPos = 10;
						}
					}
					
					// Calculate if tooltip would overflow on the right
					if (leftPos + tooltipRect.width > viewportWidth) {
						// Position so right edge of tooltip aligns with right edge of viewport minus margin
						leftPos = viewportWidth - tooltipRect.width - 10;
					}
					
					// Ensure it doesn't go off the left edge either
					if (leftPos < 10) {
						leftPos = 10;
					}
					
					tooltipBox.style.left = leftPos + 'px';
					tooltipBox.style.top = topPos + 'px';
				}, 0);
			});
			card.addEventListener('mouseleave', (e) => {
				setTimeout(() => {
					if (!tooltipBox.matches(':hover') && currentCard === card) {
						tooltipBox.style.display = 'none';
						currentCard = null;
					}
				}, 100);
			});
		});

		tooltipBox.addEventListener('mouseenter', () => {
			tooltipBox.style.display = 'block';
		});

		tooltipBox.addEventListener('mouseleave', () => {
			tooltipBox.style.display = 'none';
			currentCard = null;
		});
	});

	// Row details modal functionality
	window.closeRowDetails = function() {
		document.getElementById('rowDetailsModal').classList.remove('active');
	};

	window.showRowDetails = function(row) {
		const table = row.closest('table');
		const headers = Array.from(table.querySelectorAll('thead th'));
		const cells = Array.from(row.cells);
		
		const detailsBody = document.getElementById('rowDetailsBody');
		detailsBody.innerHTML = '';
		
		cells.forEach((cell, index) => {
			if (index < headers.length) {
				const header = headers[index];
				const columnName = header.textContent.trim();
				const cellValue = cell.textContent.trim();
				const tooltip = cell.getAttribute('title') || '';
				
				const itemDiv = document.createElement('div');
				itemDiv.className = 'row-detail-item';
				
				const labelDiv = document.createElement('div');
				labelDiv.className = 'row-detail-label';
				labelDiv.textContent = columnName;
				
				const valueDiv = document.createElement('div');
				valueDiv.className = 'row-detail-value';
				valueDiv.textContent = cellValue || '(empty)';
				
				itemDiv.appendChild(labelDiv);
				itemDiv.appendChild(valueDiv);
				
				if (tooltip) {
					const tooltipDiv = document.createElement('div');
					tooltipDiv.className = 'row-detail-tooltip';
					
					const tooltipValue = document.createElement('div');
					tooltipValue.className = 'row-detail-tooltip-value';
					tooltipValue.textContent = tooltip;
					
					tooltipDiv.appendChild(tooltipValue);
					itemDiv.appendChild(tooltipDiv);
				}
				
			detailsBody.appendChild(itemDiv);
		}
	});
	
	// Show modal and scroll to top
	const modal = document.getElementById('rowDetailsModal');
	modal.classList.add('active');
	
	// Reset scroll position after modal is visible
	setTimeout(() => {
		const modalContent = modal.querySelector('.row-details-content');
		if (modalContent) {
			modalContent.scrollTop = 0;
		}
	}, 0);
};	// Add click handlers to all table rows
	document.querySelectorAll('tbody tr').forEach(row => {
		row.addEventListener('click', function(e) {
			// Don't trigger if clicking on a link or button
			if (e.target.tagName === 'A' || e.target.tagName === 'BUTTON') {
				return;
			}
			showRowDetails(this);
		});
	});

	// Close modal when clicking outside
	document.getElementById('rowDetailsModal').addEventListener('click', function(e) {
		if (e.target === this) {
			closeRowDetails();
		}
	});
})();
</script>
<footer class="report-footer">
  <div class="creator-info">
    <p class="author-text">Author:</p>
    <div class="profile-container">
      <img src="data:image/png;base64,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" alt="Profile Picture Petri Paavola">
      <img class="black-profile" src="data:image/png;base64,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" alt="Black Profile Picture Petri Paavola">
    </div>
    <p><strong>Petri Paavola</strong><br>
      <a href="mailto:Petri.Paavola@yodamiitti.fi">Petri.Paavola@yodamiitti.fi</a><br>
      Microsoft MVP - Windows and Intune<br>
      <a href="https://intune.ninja" target="_blank">Intune.Ninja</a>
    </p>
  </div>
  <img class="company-logo" src="data:image/png;base64,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" alt="Microsoft MVP">
  <p style="margin: 0;">
    <a href="https://github.com/petripaavola/IntuneDeviceDetailsGUI" target="_blank"><strong>Download from GitHub</strong></a>
  </p>
</footer>
</body>
</html>
"@
	return $html
}

function Save-IntuneDeviceHtmlReport {
	param(
		[string]$Html,
		[string]$DeviceName
	)
	$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
	$fileSafeName = ($DeviceName -replace '[^a-zA-Z0-9_-]','_')
	$path = Join-Path -Path $script:ReportOutputFolder -ChildPath "${fileSafeName}_$timestamp.html"
	$Html | Out-File -FilePath $path -Encoding UTF8
	return $path
}

function Write-DeviceSummaryToConsole {
	param([PSObject]$Device)
	Write-Host "Creating report for device $($Device.deviceName)" -ForegroundColor Cyan
	Write-Host "OS: $($Device.operatingSystem) $($Device.osVersion)" -ForegroundColor Gray
}

function Invoke-IntuneDeviceDetailsReport {
	param(
		[string]$DeviceId,
		[switch]$ReloadCache,
		[switch]$SkipAssignments,
		[switch]$DoNotOpenReportAutomatically,
		[switch]$ExtendedReport
	)
	Initialize-IntuneSession
	if (-not $DeviceId) { return }

	$script:IntuneManagedDevice = Get-ManagedDeviceSnapshot -IntuneDeviceId $DeviceId
	if (-not $script:IntuneManagedDevice) {
		Write-Warning "Device id $DeviceId not found."
		return
	}
	Write-DeviceSummaryToConsole -Device $script:IntuneManagedDevice
	$additional = Get-AdditionalDeviceHardware -IntuneDeviceId $DeviceId
	if ($additional) {
		$additional.psobject.Properties | ForEach-Object {
			$script:IntuneManagedDevice | Add-Member -NotePropertyName $_.Name -NotePropertyValue $_.Value -Force
		}
	}

	# Add Intune device GUID to hashtable for easier access using deviceName as value
	Add-GUIDToHashtable -Object $script:IntuneManagedDevice

	$primaryContext = Get-PrimaryUserContext -Device $script:IntuneManagedDevice
	$script:PrimaryUser = $primaryContext.User
	# Add PrimaryUserGUID to hashtable for easier access
	if ($script:PrimaryUser) {
		Add-GUIDToHashtable -Object $script:PrimaryUser
	}

	$script:PrimaryUserGroupsMemberOf = $primaryContext.Groups
	# Add PrimaryUser group memberships to hashtable for easier access
	foreach ($group in $script:PrimaryUserGroupsMemberOf) {
		Add-GUIDToHashtable -Object $group
	}

	$latestContext = Get-LatestLogonContext -Device $script:IntuneManagedDevice
	# Add LatestCheckedInUserGUID to hashtable for easier access
	if ($latestContext.LatestUser) {
		Add-GUIDToHashtable -Object $latestContext.LatestUser
	}

	$script:LatestCheckedInUser = $latestContext.LatestUser
	# Add latestCheckedInUser GUID to hashtable for easier access
	if ($script:LatestCheckedInUser) {
		Add-GUIDToHashtable -Object $script:LatestCheckedInUser
	}

	$script:LatestCheckedInUserGroupsMemberOf = $latestContext.LatestGroups
	# Add LatestCheckedInUser group memberships to hashtable for easier access
	foreach ($group in $latestContext.LatestGroups) {
		Add-GUIDToHashtable -Object $group
	}

	$azureContext = Get-AzureDeviceContext -Device $script:IntuneManagedDevice
	$script:AzureADDevice = $azureContext.AzureDevice
	# Add Entra device GUID to hashtable for easier access
	if ($script:AzureADDevice) {
		Add-GUIDToHashtable -Object $script:AzureADDevice
	}

	$script:deviceGroupMemberships = $azureContext.Groups
	# Add device group memberships to hashtable for easier access
	foreach ($group in $script:deviceGroupMemberships) {
		Add-GUIDToHashtable -Object $group
	}

	$autopilot = Get-AutopilotContext -Device $script:IntuneManagedDevice
	# Add Autopilot device GUID to hashtable for easier access
	if ($autopilot -and $autopilot.Detail) {
		Add-GUIDToHashtable -Object $autopilot.Detail
	}

	$script:AutopilotDeviceWithAutpilotProfile = $autopilot.Detail

	# If Device is not Autopilot enrolled
	# Then it can be using Autopilot Device Preparation profile
	# So get Autopilot Device Preparation profile details as well if there is an enrollmentProfileName property value
	if ((-not $script:IntuneManagedDevice.autopilotEnrolled) -and ($script:IntuneManagedDevice.enrollmentProfileName)) {
		Write-Host "Detected device with Autopilot Device Preparation profile: $($script:IntuneManagedDevice.enrollmentProfileName)" -ForegroundColor Gray
		$autopilotDevicePrep = Get-AutopilotDevicePreparationContext -Device $script:IntuneManagedDevice

		if ($autopilotDevicePrep) {
			$script:AutopilotDevicePreparationPolicyWithAssignments = $autopilotDevicePrep

			Write-Host "Detected Autopilot Device Preparation profile: $($autopilotDevicePrep.displayName)" -ForegroundColor Yellow
		} else {
			$script:AutopilotDevicePreparationPolicyWithAssignments = $null
		}
	}

	$esp = Get-EnrollmentStatusPageContext -Device $script:IntuneManagedDevice
	# Add ESP profile GUID to hashtable for easier access
	if ($esp -and $esp.Detail -and $esp.Detail.enrollmentStatusPageProfile) {
		Add-GUIDToHashtable -Object $esp.Detail.enrollmentStatusPageProfile
	}
	
	$userForAssignments = if ($script:PrimaryUser) { $script:PrimaryUser.id } elseif ($script:LatestCheckedInUser) { $script:LatestCheckedInUser.id } else { '00000000-0000-0000-0000-000000000000' }
	$appAssignments = Get-ApplicationAssignmentsContext -UserId $userForAssignments -IntuneDeviceId $DeviceId -Skip:$SkipAssignments -ReloadCache:$ReloadCache

	$configPolicies = if (-not $SkipAssignments) { Get-ConfigurationPolicyReport -IntuneDeviceId $DeviceId } else { @() }
	# Add assigned configuration policy GUIDs to hashtable for easier access
	foreach ($policy in $configPolicies) {
		Add-GUIDToHashtable -Object $policy
	}

	$remediationScripts = if (-not $SkipAssignments) { 
		Get-RemediationScriptsReport -IntuneDeviceId $DeviceId `
			-DeviceGroups $script:deviceGroupMemberships `
			-PrimaryUserGroups $script:PrimaryUserGroupsMemberOf `
			-LatestUserGroups $script:LatestCheckedInUserGroupsMemberOf `
			-PrimaryUser $script:PrimaryUser `
			-LatestUser $script:LatestCheckedInUser 
	} else { @() }

	# Fetch Apple enrollment profile details if ExtendedReport is enabled and device has enrollmentProfileName
	$script:AppleEnrollmentProfileDetails = $null
	if ($ExtendedReport -and $script:IntuneManagedDevice.enrollmentProfileName -and 
		($script:IntuneManagedDevice.operatingSystem -like 'macOS*' -or $script:IntuneManagedDevice.operatingSystem -like 'iOS*')) {
		Write-Host "Fetching Apple enrollment profile details for extended report..." -ForegroundColor Cyan
		$script:AppleEnrollmentProfileDetails = Get-AppleEnrollmentProfileDetails -EnrollmentProfileName $script:IntuneManagedDevice.enrollmentProfileName
		if ($script:AppleEnrollmentProfileDetails) {
			Write-Verbose "Successfully retrieved enrollment profile: $($script:IntuneManagedDevice.enrollmentProfileName)"
		} else {
			Write-Verbose "Enrollment profile '$($script:IntuneManagedDevice.enrollmentProfileName)' not found in any DEP token"
		}
	}

	# Download script content if ExtendedReport is enabled (only for assigned scripts)
	if ($ExtendedReport -and $script:ScriptIdsToDownload -and $script:ScriptIdsToDownload.Count -gt 0) {
		Write-Host "Downloading script content for extended report... ($($script:ScriptIdsToDownload.Count) scripts tracked)" -ForegroundColor Cyan
		Write-Verbose "Script IDs to download: $($script:ScriptIdsToDownload -join ', ')"
		
		# Separate platform scripts from remediation scripts
		$platformScriptIds = @()
		$remediationScriptIds = @()
		
	foreach ($scriptId in $script:ScriptIdsToDownload) {
		# Check if it's in the platform scripts collection
		if ($script:PlatformScriptsWithAssignments | Where-Object { $_.id -eq $scriptId }) {
			Write-Verbose "Classified as platform script: $scriptId"
			$platformScriptIds += $scriptId
		}
		# Check if it's in the remediation scripts collection
		elseif ($script:RemediationScriptsWithAssignments | Where-Object { $_.id -eq $scriptId }) {
			Write-Verbose "Classified as remediation script: $scriptId"
			$remediationScriptIds += $scriptId
		}
		else {
			Write-Warning "Script ID $scriptId not found in platform or remediation collections"
		}
	}		# Download platform scripts (Windows and macOS)
		if ($platformScriptIds.Count -gt 0) {
			# Separate Windows and macOS scripts
			$windowsScriptIds = @()
			$macOSScriptIds = @()
			
			foreach ($scriptId in $platformScriptIds) {
				$platformScript = $script:PlatformScriptsWithAssignments | Where-Object { $_.id -eq $scriptId }
				if ($platformScript) {
					if ($platformScript.ScriptPlatform -eq 'Windows') {
						$windowsScriptIds += $scriptId
					}
					elseif ($platformScript.ScriptPlatform -eq 'macOS') {
						$macOSScriptIds += $scriptId
					}
				}
			}
			
			# Download Windows PowerShell scripts
			if ($windowsScriptIds.Count -gt 0) {
				Write-Host "Downloading $($windowsScriptIds.Count) Windows PowerShell platform script(s)..." -ForegroundColor Cyan
				$currentScript = 0
				
				foreach ($scriptId in $windowsScriptIds) {
					$currentScript++
					Write-Progress -Activity "Downloading PowerShell platform script content" -Status "Processing script $currentScript of $($windowsScriptIds.Count)" -PercentComplete (($currentScript / $windowsScriptIds.Count) * 100)
					
					$scriptContent = Get-PowerShellScriptContent -PowershellScriptPolicyId $scriptId
					if ($scriptContent) {
						if ($script:GUIDHashtable.ContainsKey($scriptId)) {
							$script:GUIDHashtable[$scriptId] | Add-Member -MemberType NoteProperty -Name 'scriptContentClearText' -Value $scriptContent -Force
						}
					}
				}
				Write-Progress -Activity "Downloading PowerShell platform script content" -Completed
			}
			
			# Download macOS shell scripts
			if ($macOSScriptIds.Count -gt 0) {
				Write-Host "Downloading $($macOSScriptIds.Count) macOS shell script(s)..." -ForegroundColor Cyan
				$currentScript = 0
				
				foreach ($scriptId in $macOSScriptIds) {
					$currentScript++
					Write-Progress -Activity "Downloading macOS shell script content" -Status "Processing script $currentScript of $($macOSScriptIds.Count)" -PercentComplete (($currentScript / $macOSScriptIds.Count) * 100)
					
					$scriptContent = Get-MacOSShellScriptContent -ShellScriptPolicyId $scriptId
					if ($scriptContent) {
						if ($script:GUIDHashtable.ContainsKey($scriptId)) {
							$script:GUIDHashtable[$scriptId] | Add-Member -MemberType NoteProperty -Name 'scriptContentClearText' -Value $scriptContent -Force
						}
					}
				}
				Write-Progress -Activity "Downloading macOS shell script content" -Completed
			}
		}
		
		# Download remediation scripts (both detection and remediation)
		if ($remediationScriptIds.Count -gt 0) {
			Write-Host "Downloading $($remediationScriptIds.Count) remediation script(s)..." -ForegroundColor Cyan
			$currentScript = 0
			
			foreach ($scriptId in $remediationScriptIds) {
				$currentScript++
				Write-Progress -Activity "Downloading remediation script content" -Status "Processing script $currentScript of $($remediationScriptIds.Count)" -PercentComplete (($currentScript / $remediationScriptIds.Count) * 100)
				
				# Download detection script
				$detectionContent = Get-RemediationDetectionScriptContent -ScriptPolicyId $scriptId
				if ($detectionContent) {
					if ($script:GUIDHashtable.ContainsKey($scriptId)) {
						$script:GUIDHashtable[$scriptId] | Add-Member -MemberType NoteProperty -Name 'detectionScriptContentClearText' -Value $detectionContent -Force
					}
					else {
						# Create entry if it doesn't exist
						$script:GUIDHashtable[$scriptId] = [PSCustomObject]@{
							id = $scriptId
							detectionScriptContentClearText = $detectionContent
						}
					}
				}
				
				# Download remediation script
				$remediateContent = Get-RemediationRemediateScriptContent -ScriptPolicyId $scriptId
				if ($remediateContent) {
					if ($script:GUIDHashtable.ContainsKey($scriptId)) {
						$script:GUIDHashtable[$scriptId] | Add-Member -MemberType NoteProperty -Name 'remediateScriptContentClearText' -Value $remediateContent -Force
					}
					else {
						# Create entry if it doesn't exist
						$script:GUIDHashtable[$scriptId] = [PSCustomObject]@{
							id = $scriptId
							remediateScriptContentClearText = $remediateContent
						}
					}
				}
			}
			Write-Progress -Activity "Downloading remediation script content" -Completed
		}
	}

	# Download Settings Catalog policy details if ExtendedReport is enabled (only for assigned policies)
	if ($ExtendedReport -and $script:SettingsCatalogPolicyIdsToDownload -and $script:SettingsCatalogPolicyIdsToDownload.Count -gt 0) {
		Write-Host "Downloading Settings Catalog policy details for extended report... ($($script:SettingsCatalogPolicyIdsToDownload.Count) policies tracked)" -ForegroundColor Cyan
		Write-Verbose "Settings Catalog policy IDs to download: $($script:SettingsCatalogPolicyIdsToDownload -join ', ')"
		
		$currentPolicy = 0
		foreach ($policyId in $script:SettingsCatalogPolicyIdsToDownload) {
			$currentPolicy++
			Write-Progress -Activity "Downloading Settings Catalog policy details" -Status "Processing policy $currentPolicy of $($script:SettingsCatalogPolicyIdsToDownload.Count)" -PercentComplete (($currentPolicy / $script:SettingsCatalogPolicyIdsToDownload.Count) * 100)
			
			# Download policy settings details
			$settingsData = Get-SettingsCatalogPolicyDetails -PolicyId $policyId
			if ($settingsData) {
				# Convert to readable format
				$readableSettings = ConvertTo-ReadableSettingsCatalog -SettingsData $settingsData
				
				if ($readableSettings) {
					# Store in GUIDHashtable (both readable and raw data)
					if ($script:GUIDHashtable.ContainsKey($policyId)) {
						$script:GUIDHashtable[$policyId] | Add-Member -MemberType NoteProperty -Name 'settingsCatalogDetails' -Value $readableSettings -Force
						$script:GUIDHashtable[$policyId] | Add-Member -MemberType NoteProperty -Name 'settingsRawData' -Value $settingsData -Force
					}
					else {
						# Create entry if it doesn't exist
						$script:GUIDHashtable[$policyId] = [PSCustomObject]@{
							id = $policyId
							settingsCatalogDetails = $readableSettings
							settingsRawData = $settingsData
						}
					}
					Write-Verbose "Downloaded Settings Catalog details for policy ID: $policyId"
				}
			}
		}
		Write-Progress -Activity "Downloading Settings Catalog policy details" -Completed
		
		# Analyze Settings Catalog conflicts if we have settings downloaded
		Write-Host "Analyzing Settings Catalog policies for conflicts..." -ForegroundColor Cyan
		$assignedSettingsCatalogPolicies = @()
		foreach ($policyId in $script:SettingsCatalogPolicyIdsToDownload) {
			# Get the policy object from the configuration profiles collection
			$policyObject = $Script:IntuneConfigurationProfilesWithAssignments | Where-Object { $_.id -eq $policyId }
			if ($policyObject) {
				# Settings Catalog policies use 'name' property
				$displayName = if ($policyObject.name) { $policyObject.name } elseif ($policyObject.displayName) { $policyObject.displayName } else { 'Unknown' }
				Write-Verbose "Found policy: ID=$policyId, Name='$displayName', Type=$($policyObject.'@odata.type')"
				$assignedSettingsCatalogPolicies += $policyObject
			} else {
				Write-Verbose "Policy $policyId not found in IntuneConfigurationProfilesWithAssignments collection"
			}
		}
		
		Write-Verbose "Total policies to analyze: $($assignedSettingsCatalogPolicies.Count)"
		
		$script:SettingsCatalogConflicts = Analyze-SettingsCatalogConflicts -SettingsCatalogPolicies $assignedSettingsCatalogPolicies
		
		if ($script:SettingsCatalogConflicts -and $script:SettingsCatalogConflicts.HasIssues) {
			Write-Host "Found $($script:SettingsCatalogConflicts.Conflicts.Count) conflicts and $($script:SettingsCatalogConflicts.Warnings.Count) warnings in Settings Catalog policies" -ForegroundColor Yellow
		} else {
			Write-Host "No Settings Catalog conflicts detected" -ForegroundColor Green
		}
	}
	
	# Download encrypted OMA setting values if ExtendedReport is enabled (only for assigned policies)
	if ($ExtendedReport -and $script:CustomConfigPoliciesWithSecrets -and $script:CustomConfigPoliciesWithSecrets.Count -gt 0) {
		$totalSecrets = ($script:CustomConfigPoliciesWithSecrets.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
		Write-Host "Downloading encrypted OMA setting values for extended report... ($($script:CustomConfigPoliciesWithSecrets.Count) policies, $totalSecrets secrets)" -ForegroundColor Cyan
		
		$currentSecret = 0
		foreach ($policyId in $script:CustomConfigPoliciesWithSecrets.Keys) {
			$secretIds = $script:CustomConfigPoliciesWithSecrets[$policyId]
			
			foreach ($secretId in $secretIds) {
				$currentSecret++
				Write-Progress -Activity "Downloading encrypted OMA setting values" -Status "Processing secret $currentSecret of $totalSecrets" -PercentComplete (($currentSecret / $totalSecrets) * 100)
				
				# Fetch the plain text value
				$plainTextValue = Get-OmaSettingPlainTextValue -PolicyId $policyId -SecretReferenceValueId $secretId
				
				if ($plainTextValue) {
					# Update the value in the GUIDHashtable
					if ($script:GUIDHashtable.ContainsKey($policyId)) {
						$policyObject = $script:GUIDHashtable[$policyId].Object
						if ($policyObject.omaSettings) {
							# Find and update the matching OMA setting
							foreach ($omaSetting in $policyObject.omaSettings) {
								if ($omaSetting.secretReferenceValueId -eq $secretId) {
									$omaSetting.value = $plainTextValue
									$omaSetting.isEncrypted = $false
									Write-Verbose "Updated encrypted value for policy ID: $policyId, secret: $secretId"
									break
								}
							}
						}
					}
				}
			}
		}
		Write-Progress -Activity "Downloading encrypted OMA setting values" -Completed
	}

	# DEBUG GUID Hashtable
	#$script:GUIDHashtable.GetEnumerator() | Sort-Object Name | Format-Table | Out-String | Set-Clipboard


	# ESP json has assignment groupId but no name
	# Also it may include list of "blocking" apps in ESP which are configured as applicationGUIDs
	# So try to resolve those names for better reporting and showing real names in JSON view
	# We need to run this after getting app assignments above
	$esp = Resolve-EspAssignmentGroupNames -Esp $esp

	# Autopilot configuration profile has assignment groupId but no name
	# So try to resolve those names for better reporting and showing real names in JSON view
	if ($autopilot -and $autopilot.Detail -and $autopilot.Detail.DeploymentProfileDetail) {
		$autopilot.Detail.DeploymentProfileDetail = Resolve-AssignmentGroupNames -Object $autopilot.Detail.DeploymentProfileDetail
	}


	$context = [ordered]@{
		ManagedDevice        = $script:IntuneManagedDevice
		AzureDevice          = $script:AzureADDevice
		PrimaryUser          = $primaryContext
		LatestUser           = $latestContext
		DeviceGroups         = $script:deviceGroupMemberships
		Autopilot            = $autopilot
		EnrollmentStatusPage = $esp
		Security             = $security
		AppAssignments       = $appAssignments
		ConfigurationPolicies = $configPolicies
	}

	# DEBUG $context
	#$context | ConvertTo-Json -Depth 5 | Set-Clipboard


	$html = New-IntuneDeviceHtmlReport -Context $context
	$reportPath = Save-IntuneDeviceHtmlReport -Html $html -DeviceName $script:IntuneManagedDevice.deviceName
	Write-Host ""
	Write-Host "✓ Report generated successfully!" -ForegroundColor Green
	Write-Host "  📁 $reportPath" -ForegroundColor Cyan
	Write-Host ""
	if (-not $DoNotOpenReportAutomatically) {
		Start-Process $reportPath
	}
}

## MARK: Main starts here

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " Intune Device Details GUI - HTML Report" -ForegroundColor Yellow
Write-Host " Version: $Version" -ForegroundColor Gray
Write-Host " Author: Petri Paavola / Microsoft MVP - Windows and Intune" -ForegroundColor Gray
Write-Host " GitHub: https://github.com/petripaavola/IntuneDeviceDetailsGUI" -ForegroundColor Gray
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan

Initialize-IntuneSession

if ($MyInvocation.InvocationName -ne '.') {
	$deviceIds = @()

	if ($PSBoundParameters.ContainsKey('Id') -and $Id) {
		$deviceIds += $Id
	}

	if (-not $deviceIds -and $MyInvocation.ExpectingInput) {
		foreach ($item in $input) {
			if (-not $item) { continue }
			if ($item -is [string]) {
				$deviceIds += $item
				continue
			}
			if ($item.PSObject.Properties['Id']) {
				$deviceIds += $item.Id
				continue
			}
		}
	}

	if (-not $deviceIds) {
		$resolved = Resolve-DeviceId -PipelineId $null -SearchText $SearchText
		if ($resolved) { $deviceIds += $resolved }
	}

	if (-not $deviceIds) {
		Write-Warning 'No device selected. Nothing to do.'
	}
	else {
		# Ask for report type if -ExtendedReport not specified
		$useExtendedReport = $ExtendedReport
		$useSkipAssignments = $SkipAssignments
		if (-not $PSBoundParameters.ContainsKey('ExtendedReport') -and -not $PSBoundParameters.ContainsKey('SkipAssignments') -and $deviceIds.Count -eq 1) {
			Write-Host ""
			Write-Host "📊 Report Type Selection" -ForegroundColor Yellow
			Write-Host ""
			Write-Host "  [0] Minimal Report            - Device basic info only (no assignments)" -ForegroundColor Gray
			Write-Host "  [1] Normal Report             - Basic device info, apps, configurations" -ForegroundColor White
			Write-Host "  [2] Extended Report (default) - Includes detailed policy settings, detection rules, scripts" -ForegroundColor Cyan
			Write-Host "                                  ⭐ The Hero Feature - Shows everything! (longer runtime)" -ForegroundColor Yellow
			Write-Host "                                  ⭐ Settings Catalog conflict detection!" -ForegroundColor Yellow
			Write-Host ""
			$reportChoice = Read-Host '   Select report type (0, 1 or 2, press Enter for Extended)'
			if ($reportChoice -eq '0') {
				$useExtendedReport = $false
				$useSkipAssignments = $true
				Write-Host "   ✓ Minimal Report selected" -ForegroundColor Green
			} elseif ($reportChoice -eq '1') {
				$useExtendedReport = $false
				$useSkipAssignments = $false
				Write-Host "   ✓ Normal Report selected" -ForegroundColor Green
			} else {
				$useExtendedReport = $true
				$useSkipAssignments = $false
				Write-Host "   ✓ Extended Report selected" -ForegroundColor Green
			}
			Write-Host ""
		}
		foreach ($deviceId in $deviceIds) {
			Invoke-IntuneDeviceDetailsReport -DeviceId $deviceId -ReloadCache:$ReloadCache -SkipAssignments:$useSkipAssignments -DoNotOpenReportAutomatically:$DoNotOpenReportAutomatically -ExtendedReport:$useExtendedReport
		}
	}
}

