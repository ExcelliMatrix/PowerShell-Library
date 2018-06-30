####################################################################
#                                                                  #
# Copyright (c) 2016 ExcelliMatrix, Inc. All Rights Reserved.      #
#                                                                  #
# Provided under MIT license                                       #
#                                                                  #
####################################################################

[DateTime]$ReportDateTime = [DateTime]::Now
[String]$ReportDate = $ReportDateTime.ToString("ddMMMyyyy")
[String]$ReportTime = $ReportDateTime.ToString("HHmm")
Write-Host
Write-Host
Get-DfsrBacklog -GroupName "WDS Shares" -SourceComputerName "vmStorageBZ" -DestinationComputerName "vmStroageCO" –Verbose | Format-Table -Property FullPathName | Select -First 10
Write-Host
Write-Host
Write-DfsrHealthReport -GroupName "WDS Shares" -ReferenceComputerName "vmStorageBZ" -MemberComputerName "vmStorageBZ","vmStroageCO" -Path "C:\Trash" -DomainName "Wildfire.local" -CountFiles
."C:\Program Files\Internet Explorer\iexplore.exe" "C:\Trash\Health-vmStorage-$ReportDate-$ReportTime.html"
