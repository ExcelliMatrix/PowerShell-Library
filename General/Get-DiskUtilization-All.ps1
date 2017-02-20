####################################################################
#                                                                  #
# Copyright (c) 2016-2017 ExcelliMatrix, Inc. All Rights Reserved. #
#                                                                  #
# Provided under MIT license                                       #
#                                                                  #
####################################################################

param
(
    [Parameter(Position=0)]
    [string]$ADComputerNameMask="*",

    [Parameter(Position=1)]
    [ValidateSet("None","Table")]
    [string]$OutputFormatting="Table",

    [Parameter(Position=2)]
    [ValidateSet("True","False")]
    [string]$PrettyFormatting="True",

    [Parameter(Position=3)]
    [string]$DriveLetter="", 
    
    [Parameter(Position=4)]
    [System.Decimal]$WarningThresholdPercentage=20,

    [Parameter(Position=5)]
    [System.Decimal]$CriticalThresholdPercentage=10
)

$ADComputerNames = @()
if ($ADComputerNameMask.Contains("*"))
{
    if ($PrettyFormatting -eq "True")
    {
        Write-Host
        Write-Host
        Write-Host "----- Finding AD Registered Computers (" -NoNewline
        Write-Host "$ADComputerNameMask" -NoNewline -ForegroundColor Cyan
        Write-Host ") -----" -NoNewline
    }

    $ADComputerNames = Get-ADComputer -LDAPFilter "(name=$ADComputerNameMask)" | Sort-Object Name | Foreach-Object { $_.Name }

    if ($PrettyFormatting -eq "True")
    {
        Write-Host " $($ADComputerNames.Count) matching devices in AD " -NoNewline -ForegroundColor Cyan
        Write-Host "-----"
        Write-Host
    }
}
else
{
    if ($PrettyFormatting -eq "True")
    {
        Write-Host "----- Specified Computers -----"
        $ADComputerNames += , $ADComputerNameMask
    }
}

if ($PrettyFormatting -eq "True")
{
    Write-Host "----- Retreiving Data ----- " -NoNewline
}

if ($OutputFormatting -ne "None")
{
    $ADComputerNames += ">> TOTAL <<"
}

$SortOrder = 0
$TotalDriveSize = 0
$TotalDriveFreeSpace = 0

$ReturnValue = @()
foreach ($ADComputerName in $ADComputerNames)
{
    $ComputerStatus = "Testing"
    if ($ADComputerName -eq ">> TOTAL <<")
    {
    $ADComputerUp = $true
    }
    else
    {
      $ADComputerUp = Test-Connection -BufferSize 32 -Count 1 -ComputerName $ADComputerName -Quiet
    }

    $SystemDriveLetters = @()
    if ($DriveLetter -ne "")
    {
        $SystemDriveLetters += $DriveLetter
    }
    else
    {
        $SystemDriveLetters += "C:"
        $SystemDriveLetters += "D:"
        $SystemDriveLetters += "E:"
        $SystemDriveLetters += "F:"
        $SystemDriveLetters += "G:"
    }

    $MeasureCommand = Measure-Command {
        foreach ($SystemDriveLetter in $SystemDriveLetters)
        {
            if (($ComputerStatus -ne "Inaccessible") -and ($ComputerStatus -ne "----------------"))
            {
                $SortOrder = $SortOrder + 1
                $ReturnItem = "" | Select SortOrder, ComputerName, ComputerStatus, DriveLetter, TotalBytes, UsedBytes, FreeBytes, FreePercentage, FreeRank, Concern, FreeSize
                $ReturnItem.SortOrder    = $SortOrder
                $ReturnItem.ComputerName = $ADComputerName
                $ReturnItem.FreeSize     = 0

                if ($ADComputerUp -eq $true)
                {
                    if ($ADComputerName -ne ">> TOTAL <<")
                    {
                        $Drive = Get-WmiObject Win32_LogicalDisk -ComputerName $ADComputerName -Filter "DeviceID='$SystemDriveLetter'"
                        $DriveSize = $Drive.Size
                        $DriveFreeSpace = $Drive.FreeSpace

                        $TotalDriveSize += $DriveSize
                        $TotalDriveFreeSpace += $DriveFreeSpace
                    }
                    else
                    {
                        $DriveSize = $TotalDriveSize
                        $DriveFreeSpace = $TotalDriveFreeSpace
                    }

                    if ($DriveSize -gt 0)
                    {
                        $ComputerStatus = "Available"

                        $DriveUsedSpace = $DriveSize - $DriveFreeSpace
                        $DrivePercentFree = [Math]::Round(($DriveFreeSpace / $DriveSize) * 100, 2)

                        $ReturnItem.DriveLetter    = $SystemDriveLetter
                        if ($PrettyFormatting -eq "True")
                        {
                            $PrettyDriveSize           = "$([Math]::Round($DriveSize / (1024 * 1024 * 1024), 2) + 0.001)xx" -replace "1xx", " GB"
                            $PrettyDriveUsedSpace      = "$([Math]::Round($DriveUsedSpace / (1024 * 1024 * 1024), 2) + 0.001)xx" -replace "1xx", " GB"
                            $PrettyDriveFreeSpace      = "$([Math]::Round($DriveFreeSpace / (1024 * 1024 * 1024), 2) + 0.001)xx" -replace "1xx", " GB"
                            $PrettyDrivePercentageFree = "$($DrivePercentFree + 0.001)xx" -replace "1xx", "%"

                            $ReturnItem.TotalBytes     = $PrettyDriveSize
                            $ReturnItem.UsedBytes      = $PrettyDriveUsedSpace
                            $ReturnItem.FreeBytes      = $PrettyDriveFreeSpace
                            $ReturnItem.FreePercentage = $PrettyDrivePercentageFree
                            $ReturnItem.FreeSize       = $DriveFreeSpace
                        }
                        else
                        {
                            $ReturnItem.TotalBytes     = [Math]::Round($DriveSize/ (1024 * 1024 * 1024), 2)
                            $ReturnItem.UsedBytes      = [Math]::Round($DriveUsedSpace/ (1024 * 1024 * 1024), 2)
                            $ReturnItem.FreeBytes      = [Math]::Round($DriveFreeSpace/ (1024 * 1024 * 1024), 2)
                            $ReturnItem.FreePercentage = $DrivePercentFree
                        }

                        if (($DrivePercentFree -le $WarningThresholdPercentage) -or ($DrivePercentFree -le $cr))
                        {
                            if ($DrivePercentFree -le $WarningThresholdPercentage)
                            {
                                $ReturnItem.Concern = "Low Disk Space"
                            }
                            else
                            {
                                $ReturnItem.Concern = "Low Disk Space - Critical"
                            }
                        }

                        if ($ADComputerName -eq ">> TOTAL <<")
                        {
                            $ComputerStatus = "----------------"
                            $ReturnItem.DriveLetter = "--"
                            $ReturnItem.SortOrder = 999999
                        }

                        $ReturnItem.ComputerStatus = $ComputerStatus
                        $ReturnValue += $ReturnItem
                    }
                }
                else
                {
                    $ComputerStatus = "Inaccessible"
                    $ReturnItem.ComputerStatus = $ComputerStatus
                    $ReturnValue += $ReturnItem
                }
            }
        }
    }

    if ($PrettyFormatting -eq "True")
    {
        if ($ComputerStatus -eq "Inaccessible")
        {
            Write-Host "+" -NoNewline -BackgroundColor Yellow -ForegroundColor Red
        }
        else
        {
            #Write-Host "($($MeasureCommand.TotalMilliseconds))" -NoNewline
            if ($MeasureCommand.TotalMilliseconds -ge 250)
            {
                Write-Host "+" -NoNewline -BackgroundColor Yellow
            }
            else
            {
                Write-Host "+" -NoNewline
            }
        }
    }
}

$FreeRank = -1
$ReturnValueTemp = @()
$ReturnValue | Sort-Object FreeSize -Descending | foreach {
    $FreeRank = $FreeRank + 1
    $_.FreeRank = $FreeRank
    $ReturnValueTemp += $_
}
$ReturnValue = $ReturnValueTemp | Sort-Object SortOrder

if ($PrettyFormatting -eq "True")
{
    Write-Host
}

if ($OutputFormatting -eq "None")
{
    $ReturnValue
}
else
{
    $ReturnValue | Format-Table @{n=' ComputerName ';e={$_.ComputerName};align='left'},
        @{n=' ComputerStatus ';e={$_.ComputerStatus};align='center'},
        @{n=' DriveLetter ';e={$_.DriveLetter};align='center'},
        @{n='  TotalBytes  ';e={$_.TotalBytes};align='right'},
        @{n='  UsedBytes  ';e={$_.UsedBytes};align='right'},
        @{n='  FreeBytes  ';e={$_.FreeBytes};align='right'},
        @{n=' FreePercentage ';e={$_.FreePercentage};align='right'},
        @{n=' FreeRank ';e={$_.FreeRank};align='center'},
        @{n=' Concern ';e={$_.Concern};align='left'} -AutoSize
}
if ($PrettyFormatting -eq "True")
{
    Write-Host
    Write-Host
}
