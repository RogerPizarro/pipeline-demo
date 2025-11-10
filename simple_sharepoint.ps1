# Export users with access to a SharePoint site
Connect-SPOService -Url "https://tenant-admin.sharepoint.com"
Get-SPOUser -Site "https://tenant.sharepoint.com/sites/Example" |
  Select LoginName, Email, IsSiteAdmin |
  Export-Csv "C:\Reports\SiteUsers.csv" -NoTypeInformation
