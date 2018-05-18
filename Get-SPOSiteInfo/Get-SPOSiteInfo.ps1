    <#
    .Synopsis
        Retrieve SharePoint Online information. 
    .DESCRIPTION
        This function will retrieve SharePoint Online information like Site Collections, owners, storage used, etc.
        Report (.xlsx file) will be generated with color-coded information about Site Collections connected to Office 365 Groups, Sites in Read-Only mode, and when storage used has reached 80%.
    .EXAMPLE
        C:\PS>  Get-SPOSiteInfo -TenantName Contoso 
    .EXAMPLE
        C:\PS>  Get-SPOSiteInfo (if -TenantName parameter is not entered, user will be prompted for it)
    .INPUTS
        None
    .OUTPUTS
        System.Object
    .NOTES
        None
    .COMPONENT
        This function also uses SharePoint PowerShell PnP (Patterns & Practices) available on GitHub: https://github.com/SharePoint/PnP-PowerShell
        Therefore, SharePoint PowerShell PnP is a requirement for the function to work.
    .LINK
    SharePoint PnP PowerShell (Patterns & Practices) available on GitHub: https://github.com/SharePoint/PnP-PowerShell
    #>
function Get-SPOSiteInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Office 365 Tenant name", Position = 0)]
        [string]$TenantName
    )
    #Report placed onto the user's desktop
    $reportLocation = "C:\Users\$env:USERNAME\Desktop\"
    $startTime = Get-Date -Format MM-dd-yyyy-HH-mm-ss
    $SPODataFile = $reportLocation + "SPOSitesReport_$TenantName _$startTime.xlsx"
    
    # Connect to SPO
    $cred = Get-Credential
    try {
        Connect-PnPOnline -Url "https://$TenantName-admin.sharepoint.com" -Credentials $cred -ErrorAction Stop
        Write-Host "You are now connected to SPO." -ForegroundColor Green
    }
    catch {
        Write-Error -Message "Credentials are not correct. Please try again."
        break
    }

    #Get the Site Collections information
    Write-Host "Retrieving the information. Be patient..." -ForegroundColor Yellow
    $SPOData = Get-PnPTenantSite | select-object Title, Url, Template, Owner, `
    @{n = 'Lock State'; e = { ($_.LockState) }}, `
    @{n = 'Storage Used (GB)'; e = { (($_.StorageUsage) / 1024).ToString("N")}}, `
    @{n = 'Storage Limit (GB)'; e = { (($_.StorageMaximumLevel) / 1024).ToString("N") }}, `
    @{n = 'Storage % Used'; e = { (($_.StorageUsage) / ($_.StorageMaximumLevel))}}, `
    @{n = 'Storage Warning (GB)'; e = { ($_.StorageWarningLevel) / 1024 -as [int] }}, `
    @{n = 'Server Resource Quota'; e = { ($_.UserCodeMaximumLevel) }}, `
    @{n = 'Server Resource Warning (at)'; e = { ($_.UserCodeWarningLevel) }}


    #Variable for formatting 
    $cfReadOnlyText = New-ConditionalText -Text 'ReadOnly' -BackgroundColor Orange -ConditionalTextColor White 
    $cfGroupText = New-ConditionalText -Text 'GROUP#0' -BackgroundColor Wheat
    
    #Creating the workbook in a variable called "$myWorkbook"
    $myWorkbook = $SPOData | Export-Excel -Path $SPODataFile -WorkSheetname "SPOData" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow -ConditionalText $cfReadOnlyText, $cfGroupText -PassThru
    
    #Formatting the "Storage % Used" with RED bground if >= 80%
    $SPODataWS = $myWorkbook.Workbook.Worksheets["SPOData"]
    Set-Format -WorkSheet $SPODataWS -Range "H2:H500000" -NumberFormat "0.00%" -AutoFit
    Add-ConditionalFormatting -WorkSheet $SPODataWS -Range "H2:H500000" -RuleType GreaterThanOrEqual -ConditionValue '80.00%' -ForeGroundColor White -BackgroundColor "Red"
    Set-CellStyle -WorkSheet $SPODataWS -LastColumn 1 -Pattern Solid -Color LightGray

    #Exporting the data
    Export-Excel -ExcelPackage $myWorkbook -Show 

    Write-Host "Report created : $SPODataFile" -ForegroundColor Green

} #end of the function

