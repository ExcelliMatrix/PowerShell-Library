####################################################################
#                                                                  #
# Copyright (c) 2016 ExcelliMatrix, Inc. All Rights Reserved.      #
#                                                                  #
# Provided under MIT license                                       #
#                                                                  #
####################################################################

param
(
  [string]$ComputerName = "",
  [bool]$ShowDetail=$false
)

Write-Host
Write-Host

Write-Host "$(Get-Date)"

Add-Type @"
  using System;
  public class DFSRStateItemType
  {
    public string ComputerName;
    public string FileName;
    public string UpdateState;
    public bool   Inbound;
    public string SourceComputerName;
    public string Path;
    public long   FileSize;
  }
"@

$ComputerNames = @()
if ($ComputerName -eq "")
{
  $ReplicationGroups = Get-DfsReplicationGroup -IncludeSysvol
  foreach ($ReplicationGroup in $ReplicationGroups)
  {
    $ReplicationGroupMembers = Get-DfsrMembership -GroupName "$($ReplicationGroup.GroupName)" | Select-Object -Property ComputerName | Sort-Object -Property ComputerName -Unique
    foreach ($ReplicationGroupMember in $ReplicationGroupMembers)
    {
      $ComputerNames += $ReplicationGroupMember.ComputerName
    }
  }
}
else
{
  $ComputerNames += $ComputerName
}

$DFSRStateItems = @()
foreach ($ComputerName in $ComputerNames)
{
  $Results = Get-DFSRState -ComputerName "$($ComputerName)"
  foreach ($Result in  $Results)
  {
    $DFSRStateItem = New-Object DFSRStateItemType
    $DFSRStateItem.ComputerName       = $ComputerName
    $DFSRStateItem.FileName           = $Result.FileName
    $DFSRStateItem.UpdateState        = $Result.UpdateState
    $DFSRStateItem.Inbound            = $Result.Inbound
    $DFSRStateItem.SourceComputerName = $Result.SourceComputerName
    $DFSRStateItem.Path               = $Result.Path
    $DFSRStateItem.FileSize           = 0

    if ($Result.Path -ne "")
    {
      $SourceComputerName1 = $ComputerName
      $SourceComputerName2 = $Result.SourceComputerName

      $FileName  = "$($DFSRStateItem.Path)"
      $FileName  = "$($FileName.Replace('C:', 'C$'))"
      $FileName  = "$($FileName.Replace('D:', 'D$'))"
      $FileName  = "$($FileName.Replace('E:', 'E$'))"
      $FileName  = "$($FileName.Replace('F:', 'F$'))"
      $FileName  = "$($FileName.Replace('G:', 'G$'))"
      $FileName  = "$($FileName.Replace('H:', 'H$'))"

      $FileName1 = "\\$SourceComputerName1\$FileName"
      $FileName2 = "\\$SourceComputerName2\$FileName"

      $FileSize = 0

      if([System.IO.File]::Exists($FileName1))
      {
        $FileSize = (Get-Item "$FileName1").Length
      }
      else
      {
        if([System.IO.File]::Exists($FileName2))
        {
          $FileSize = (Get-Item "$FileName2").Length
        }
      }

      $DFSRStateItem.FileSize = $FileSize
    }

    $DFSRStateItems += $DFSRStateItem
  }
}

if ($ShowDetail -eq $true)
{
  $DFSRStateItems | Sort-Object Inbound, UpdateState -Descending | Format-Table ComputerName, FileName, FileSize, UpdateState, Inbound, SourceComputerName, Path -auto
}

$DFSRStateItems | Group-Object ComputerName, UpdateState | Select Name, Count | Sort-Object Name | Format-Table -AutoSize
