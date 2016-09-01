#------------------------------------------------------------------------------------------- 
# Name:            Get-SPFarmSitesInfo
# Description:     This script will export SharePoint 2013 SPFarm all site Title,Url,MaxStorage,Storageused,SiteQuotaUsed,TotalHits,TotalUniqueUsers  data to a CSV file
# Usage:           Run the function with the required parameters.  
#                  Scope can be all SPSite and / or all SPWeb objects in a Web Application
#
#                  Get-SPFarmSitesInfo  -OutputFilePath "C:\SPFarmAllSiteInfoData.csv" 
#  
# Author:          6Qiang
#
# Reference:       SearchServiceApplicationProxy.GetRollupAnalyticsItemData method parameters
#                  http://msdn.microsoft.com/en-us/library/office/microsoft.office.server.search.administration.searchserviceapplicationproxy.getrollupanalyticsitemdata.aspx
#
# Inspiration:     http://www.sharepointtalk.net/2014/02/query-sharepoint-search-analytics-using.html
# 				   https://sp2013wade.codeplex.com/
#                  http://gallery.technet.microsoft.com/office/Get-SharePoint-Web-19cd2137 (Ivan Josipovic)
#------------------------------------------------------------------------------------------- 


asnp microsoft.sharepoint.powershell

function Get-SPFarmSitesInfo {
	Param(
		[string]$OutputFilePath
	)

	# Delete CSV file if existing
	If (Test-Path $OutputFilePath) {
		Remove-Item $OutputFilePath
	}

	$infoColl = @()

	ForEach($s in Get-SPSite) { 
		$web = $s.RootWeb;
		$siteTitle = $web.Title
		$siteUsage = $s.Usage;

		#站点分配的额度
		if ($s.Quota.Storagemaximumlevel -gt 0) 
		{
			[int]$MaxStorage = $s.Quota.StorageMaximumLevel /1MB
		} 
		else {
			$MaxStorage = "0"
		}; 
		#可用额度
		if ($s.Usage.Storage -gt 0) {
			[int]$StorageUsed = $s.Usage.Storage /1MB
		};
		#额度使用率
		if ($Storageused -gt 0 -and $Maxstorage -gt 0){
			[int]$SiteQuotaUsed = $Storageused / $Maxstorage * 100
		} else {
			$SiteQuotaUsed = "0"
		}; 
		$search = get-spenterpriseSearchserviceApplication
		$usage = $search.GetRollupAnalyticsItemData(1,[System.Guid]::Empty,$s.ID,[system.guid]::Empty)
		#总点击量PV
		$TotalHits = $usage.TotalHits
		#总独立访客UV
		$TotalUniqueUsers = $usage.TotalUniqueUsers 

		$infoObject = New-Object PSObject 
		Add-Member -inputObject $infoObject -memberType NoteProperty -name "SiteTitle" -value $web.Title
		Add-Member -inputObject $infoObject -memberType NoteProperty -name "SiteUrl" -value $web.Url 
		Add-Member -inputObject $infoObject -memberType NoteProperty -name "MaxStorage" -value $MaxStorage
		Add-Member -inputObject $infoObject -memberType NoteProperty -name "Storageused" -value $Storageused
		Add-Member -inputObject $infoObject -memberType NoteProperty -name "SiteQuotaUsed" -value $SiteQuotaUsed
		Add-Member -inputObject $infoObject -memberType NoteProperty -name "TotalHits" -value $TotalHits 
		Add-Member -inputObject $infoObject -memberType NoteProperty -name "TotalUniqueUsers" -value $TotalUniqueUsers
		$infoColl += $infoObject
	}

	$infoColl | Export-Csv -path $OutputFilePath -Encoding UTF8 -NoTypeInformation 
}

