####################################################################
#                                                                  #
# Copyright (c) 2016 ExcelliMatrix, Inc. All Rights Reserved.      #
#                                                                  #
# Provided under MIT license                                       #
#                                                                  #
####################################################################

param
(
  [string] $GroupName = "",
  [string] $DomainName = "",
  [string] $SourceComputerName = "",
  [string] $DestinationComputerName = "",
  [string] $OutputPath = "C:\Trash"
)

[DateTime]$ReportDateTime = [DateTime]::Now
[string]$ReportDate = $ReportDateTime.ToString("ddMMMyyyy")
[string]$ReportTime = $ReportDateTime.ToString("HHmm")
[string]$FileName = "$($OutputPath)\DFSReports\Health-$($GroupName)-$ReportDate-$ReportTime.html"

Write-Host
Write-Host
Get-DfsrBacklog -GroupName "$GroupName" -SourceComputerName "$SourceComputerName" -DestinationComputerName "$DestinationComputerName" –Verbose | Format-Table -Property FullPathName | Select -First 10

Write-Host
Write-Host
Write-Host "Generating file '$FileName'"
Write-DfsrHealthReport -GroupName "$GroupName" -ReferenceComputerName "$SourceComputerName" -MemberComputerName "$SourceComputerName","$DestinationComputerName" -Path "$OutputPath" -DomainName "$DomainName" -CountFiles

#."C:\Program Files\Internet Explorer\iexplore.exe" "$FileName"
