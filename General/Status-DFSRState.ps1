####################################################################
#                                                                  #
# Copyright (c) 2016 ExcelliMatrix, Inc. All Rights Reserved.      #
#                                                                  #
# Provided under MIT license                                       #
#                                                                  #
####################################################################

param
(
  [string]$ComputerName = "$env:computername",
  [bool]$ShowDetail=$false,
  [bool]$Verbose=$false
)

Write-Host
Write-Host

Write-Host "$(Get-Date)"
Get-DFSRState -ComputerName "$($ComputerName)" | Sort-Object Inbound, UpdateState -Descending | Format-Table FileName, UpdateState, Inbound, SourceComputerName, Path -auto
