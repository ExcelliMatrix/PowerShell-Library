####################################################################
#                                                                  #
# Copyright (c) 2016 ExcelliMatrix, Inc. All Rights Reserved.      #
#                                                                  #
# Provided under MIT license                                       #
#                                                                  #
####################################################################

param
(
  [string]$SourceComputer = "*",
  [string]$DestinationComputer = "*",
  [bool]$ShowDetail=$false,
  [bool]$Verbose=$false
)

Write-Host
if ($Verbose -eq $true)
{
  Write-Host "Retreiving List of Replication Groups..." -NoNewline
}
$ReplicationGroups = Get-DfsReplicationGroup -IncludeSysvol
if ($Verbose -eq $true)
{
  Write-Host "$($ReplicationGroups.Count) Groups Found..." -NoNewline
  Write-Host "Done"
}

Add-Type @"
  using System;
  public class BacklogItemType
  {
    public string SourceComputerName;
    public string DestinationComputerName;
    public string FileName;
    public string FullPathName;
    public DateTime CreateTime;
    public DateTime UpdateTime;
    public TimeSpan BacklogAge
    {
      get
      {
        return (DateTime.Now - UpdateTime);
      }
    }
    public int BacklogDays
    {
      get
      {
        return (DateTime.Now - UpdateTime).Days;
      }
    }
    public string Summary
    {
      get
      {
        string sReturnValue = "";
        if (BacklogAge.TotalMinutes > 30)
          sReturnValue += "* Excessive Time In Backlog *";
        return sReturnValue;
      }
    }
  }
"@

$BacklogItems = @()
Write-Host "     $(Get-Date)"
foreach ($ReplicationGroup in $ReplicationGroups)
{
  if ($Verbose -eq $true)
  {
    Write-Host "     Replication Group: ""$($ReplicationGroup.GroupName)""..." -NoNewline
  }
  $ReplicationGroupMembers = Get-DfsrMembership -GroupName "$($ReplicationGroup.GroupName)" | Select-Object -Property ComputerName | Sort-Object -Property ComputerName -Unique
  if ($Verbose -eq $true)
  {
    Write-Host "$($ReplicationGroupMembers.Count) Members Found."
    foreach ($ReplicationGroupMember in $ReplicationGroupMembers)
    {
      Write-Host "          ""$($ReplicationGroupMember.ComputerName)"""
    }

    Write-Host
    Write-Host
  }

  if ($Verbose -eq $true)
  {
    Write-Host "SourceComputer = '$SourceComputer'"
    Write-Host "DestinationComputer = '$DestinationComputer'"
  }

  foreach ($ReplicationGroupMember_Outer in $ReplicationGroupMembers)
  {
    if (($SourceComputer -eq "*") -or ($ReplicationGroupMember_Outer.ComputerName.ToLower() -eq $SourceComputer.ToLower()))
    {
      foreach ($ReplicationGroupMember_Inner in $ReplicationGroupMembers)
      {
        if ($ReplicationGroupMember_Outer.ComputerName -ne $ReplicationGroupMember_Inner.ComputerName)
        {
          if (($DestinationComputer -eq "*") -or ($ReplicationGroupMember_Inner.ComputerName.ToLower() -eq $DestinationComputer.ToLower()))
          {
            if ($Verbose -eq $true)
            {
              Write-Host "Get-DfsrBacklog -GroupName $($ReplicationGroup.GroupName) -SourceComputerName $($ReplicationGroupMember_Outer.ComputerName) -DestinationComputerName $($ReplicationGroupMember_Inner.ComputerName)"
            }
            Write-Host "     Checking Backlog from $($ReplicationGroupMember_Outer.ComputerName) to $($ReplicationGroupMember_Inner.ComputerName)..." -NoNewline
            $Backlogs = Get-DfsrBacklog -GroupName "$($ReplicationGroup.GroupName)" -SourceComputerName "$($ReplicationGroupMember_Outer.ComputerName)" -DestinationComputerName "$($ReplicationGroupMember_Inner.ComputerName)" -ErrorAction SilentlyContinue
            Write-Host "$($Backlogs.Count) Backlog Items Found." -ForegroundColor Yellow
            foreach ($Backlog in $Backlogs)
            {
              $BacklogItem = New-Object BacklogItemType
              $BacklogItem.SourceComputerName = "$($ReplicationGroupMember_Outer.ComputerName)"
              $BacklogItem.DestinationComputerName = "$($ReplicationGroupMember_Inner.ComputerName)"
              $BacklogItem.FileName = $Backlog.FileName
              $BacklogItem.FullPathName = $Backlog.FullPathName
              $BacklogItem.CreateTime = $Backlog.CreateTime
              $BacklogItem.UpdateTime = $Backlog.UpdateTime
            
              $BacklogItems += $BacklogItem
            }
          }
        }
      }
    }
  }
}

if ($BacklogItems.Count -eq 0)
{
  Write-Host
  Write-Host "No Report Generated" -ForegroundColor Cyan
  Write-Host
}
else
{
  if ($Verbose -eq $true)
  {
    Write-Host "     $($BacklogItems.Count) Backlog Entries" -ForegroundColor Yellow
  }
  if ($ShowDetail -eq $false)
  {
    $BacklogItems | Sort-Object BacklogDays | Group-Object -Property BacklogDays | Select-Object Name, Count | Format-Table @{Label="Backlog Days"; Expression={$_.Name}; Align='Right'}, Count

    Write-Host
    Write-Host "Detailed output surppressed." -ForegroundColor Cyan
    Write-Host
  }
  else
  {
    $BacklogItems | Sort-Object -Property SourceComputerName, DestinationComputerName, BacklogAge, FileName |
      Format-Table -AutoSize `
        @{Label="Source"; Expression={ $($_.SourceComputerName)}; align='left' },
        @{Label="Destination"; Expression={ $($_.DestinationComputerName)}; align='left' },
        @{Label="Backlog Age"; Expression={ $($_.BacklogAge.ToString("dd\.hh\:mm\:ss")) }; align='left' },
        @{Label="File"; Expression={ $($_.FileName)}; align='left' },
        @{Label="Create Time/Time"; Expression={ $($_.CreateTime.ToString("MM-dd-yyyy HH:mm:ss")) }; align='left' },
        @{Label="Update Date/Time"; Expression={ $($_.UpdateTime.ToString("MM-dd-yyyy HH:mm:ss")) }; align='left' },
        @{Label="Summary"; Expression={ $($_.Summary)}; align='left' },
        @{Label="Full Path"; Expression={ $($_.FullPathName)}; align='left' }

    $BacklogItems | Sort-Object BacklogDays | Group-Object -Property BacklogDays | Select-Object Name, Count | Format-Table @{Label="Backlog Days"; Expression={$_.Name}; Align='Right'}, Count
  }
  Write-Host "Total Backlog Entries Found: $($BacklogItems.Count)" -ForegroundColor Yellow
  Write-Host
  Write-Host
}