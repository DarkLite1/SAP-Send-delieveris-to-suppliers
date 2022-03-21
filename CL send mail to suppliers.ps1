#Requires -Version 5.1
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Send a mail to the suppliers about their deliveries the day before.

.DESCRIPTION
    SAP generates a .ASC file that contains the deliveries of the previous day. 
    This file is used to calculate transport costs by the suppliers.

    The file created on the day that the script executes is the one that is 
    converted to an Excel file and send to the supplier by mail.

    In case there is no .ASC file created on the day that the script runs, 
    nothing is done and no mail is sent out.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\CL\$ScriptName",
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        #region Logging
        try {
            $LogParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $LogFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion
        
        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
        
        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion
        
        #region Test .json file properties
        if (-not ($MailTo = $file.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }
        if (-not ($Suppliers = $file.Suppliers)) {
            throw "Input file '$ImportFile': No 'Suppliers' found."
        }
        foreach ($s in $Suppliers) {
            #region Path
            if (-not $s.Path) {
                throw "Input file '$ImportFile': No 'Path' found in one of the 'Suppliers'."
            }
            if (-not ($s.Path -Match '^\\\\')) {
                throw "Input file '$ImportFile': 'Path' needs to be a UNC path"
            }
            if (-not (Test-Path -LiteralPath $s.Path -PathType Container)) {
                throw "Input file '$ImportFile': 'Path' folder '$($s.Path)' not found"
            }
            #endregion
        }
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    try {
        $ascFiles = Get-ChildItem $Path -Filter '*.ASC' -File

        $convertedAscFiles = foreach ($file in $ascFiles) {
            $fileContent = Get-Content -LiteralPath $file.FullName

            $data = foreach ($line in $fileContent) {
                [PSCustomObject]@{
                    Plant               = $line.SubString(0, 4).Trim()
                    ShipmentNumber      = [int]$line.SubString(4, 10).Trim()
                    DeliveryNumber      = [int]$line.SubString(14, 30).Trim()
                    ShipToNumber        = [int]$line.SubString(44, 10).Trim()
                    ShipToName          = $line.SubString(54, 35).Trim()
                    Address             = $line.SubString(89, 35).Trim()
                    City                = $line.SubString(124, 35).Trim()
                    MaterialNumber      = [int]$line.SubString(159, 18).Trim()
                    MaterialDescription = $line.SubString(177, 40).Trim()
                    Tonnage             = $line.SubString(217, 6).Trim()
                    LoadingDate         = $(
                        if ($loadingDate = $line.SubString(223, 8).Trim()) {
                            [DateTime]::ParseExact($loadingDate, 'yyyyMMdd', $null)
                        }
                    )
                    DeliveryDate        = $(
                        $deliveryDate = $line.SubString(231, 8).Trim()
                        $deliveryTime = $line.SubString(239, 6).Trim()
                        if ($deliveryDate -and $deliveryTime) {
                            [DateTime]::ParseExact(
                    ($deliveryDate + $deliveryTime), 'yyyyMMddHHmmss', $null
                            )
                        }
                        elseif ($deliveryDate) {
                            [DateTime]::ParseExact($deliveryDate, 'yyyyMMdd', $null)
                        }
                    )
                    TruckID             = $line.SubString(245, 20).Trim()
                    PickingStatus       = $line.SubString(265, 1).Trim()
                    SiloBulkID          = $line.SubString(266, ($line.Length - 266)).Trim()
                }
            }

            [PSCustomObject]@{
                Name    = $file.BaseName
                Content = $data
            }
        }

        foreach (
            $convertedFile in 
            $convertedAscFiles | Where-Object { $_.Content }
        ) {
            $excelParams = @{
                Path          = Join-Path $logFolder ($convertedFile.Name + '.xlsx')
                WorksheetName = 'Data'
                TableName     = 'Data'
                FreezeTopRow  = $true
                AutoSize      = $true
            }
            $convertedFile.Content | Export-Excel @excelParams
        }
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    try {
        
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}