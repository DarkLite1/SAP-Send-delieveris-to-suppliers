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
    [String]$Mailbox = 'BNL.ServiceDesk@heidelbergcement.com',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\CL\$ScriptName",
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Function Test-ValidEmailAddress { 
        Param(
            [Parameter(Mandatory)]
            [String]$EmailAddress
        )
        try {
            $null = [MailAddress]$EmailAddress
            return $true
        }
        catch {
            return $false
        }
    }

    try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
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
        if (-not ($MailFrom = $file.MailFrom)) {
            throw "Input file '$ImportFile': No 'MailFrom' addresses found."
        }
        if (-not ($Suppliers = $file.Suppliers)) {
            throw "Input file '$ImportFile': No 'Suppliers' found."
        }
        foreach ($s in $Suppliers) {
            #region Name
            if (-not $s.Name) {
                throw "Input file '$ImportFile': Property 'Name' is missing in 'Suppliers'."
            }
            #endregion

            #region Path
            if (-not $s.Path) {
                throw "Input file '$ImportFile': Property 'Path' is missing in 'Suppliers' for '$($s.Name)'."
            }
            if (-not (Test-Path -LiteralPath $s.Path -PathType Container)) {
                throw "Input file '$ImportFile': 'Path' folder '$($s.Path)' not found for '$($s.Name)'"
            }
            #endregion
            
            #region MailTo
            if (-not $s.MailTo) {
                throw "Input file '$ImportFile': Property 'MailTo' is missing in 'Suppliers' for '$($s.Name)'."
            }
            if (-not (Test-ValidEmailAddress -EmailAddress $s.MailTo)) {
                throw "Input file '$ImportFile': 'MailTo' value '$($s.MailTo)' is not a valid e-mail address for supplier '$($s.Name)'."
            }
            #endregion

            #region NewerThanDays
            if ($s.PSObject.Properties.Name -notContains 'NewerThanDays') {
                throw "Input file '$ImportFile': Property 'NewerThanDays' is missing for supplier '$($s.Name)'. Use number '0' to only handles files with creation date today."
            }
            if (-not ($s.NewerThanDays -is [int])) {
                throw "Input file '$ImportFile': 'NewerThanDays' needs to be a number, the value '$($s.NewerThanDays)' is not supported. Use number '0' to only handle files with creation date today."
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
        $mailParams = @{
            Service        = New-EwsServiceHC
            From           = $MailFrom
            SentItemsPath  = ('\PowerShell\' + $ScriptName + ' SENT')
            EventLogSource = $ScriptName
            ErrorAction    = 'Stop'
            Body           = 'test'
        }
        
        #region Create mailbox folders
        $mailFolderParams = @{
            Mailbox = $mailParams.From
            Path    = $mailParams.SentItemsPath
            Service = $mailParams.Service
        }
        New-MailboxFolderHC @mailFolderParams
        #endregion

        foreach ($s in $Suppliers) {
            $mailParams.Attachments = @()
            
            #region Get .ASC files
            $compareDate = (Get-Date).addDays(-$s.NewerThanDays)
            $M = "Get .ASC files for supplier '$($s.Name)'"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $getParams = @{
                LiteralPath = $s.Path
                Filter      = '*.ASC' 
                File        = $true
            }
            $ascFiles = Get-ChildItem @getParams |
            Where-Object { $_.CreationTime.Date -ge $compareDate.Date }

            $M = "Found $($ascFiles.Count) .ASC files for supplier '$($s.Name)' older than '$($compareDate.Date)'"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            #endregion

            #region Convert files to objects
            $exportToExcel = foreach ($file in $ascFiles) {
                $M = "Convert file '$($file.FullName)' to objects for Excel"
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
                $fileContent = Get-Content -LiteralPath $file.FullName
                
                foreach ($line in $fileContent) {
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
                        File                = $file.BaseName
                    }
                }
                
                $mailParams.Attachments += $file.FullName
            }
            #endregion

            if ($exportToExcel) {
                #region Export to Excel
                $M = "Export '$($exportToExcel.Count)' objects to Excel"
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
                
                $logParams.Name = $s.Name

                $excelParams = @{
                    Path          = "$(New-LogFileNameHC @logParams).xlsx"
                    WorksheetName = 'Data'
                    TableName     = 'Data'
                    FreezeTopRow  = $true
                    AutoSize      = $true
                }
                $exportToExcel | Export-Excel @excelParams

                $M = "Exported '$($exportToExcel.Count)' rows to Excel file '$($excelParams.Path)'"
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

                $mailParams.Attachments += $excelParams.Path
                #endregion

                #region Send mail to end user
                $mailParams.To = $s.MailTo
                $mailParams.Bcc = @($ScriptAdmin)
                if ($s.MailBcc) { 
                    $mailParams.Bcc += $s.MailBcc 
                }
                $mailParams.Body = '<p>Dear</p><p>Since <b>{0}</b> there {1}.</p><p><i>Check the attachments for details.</i></p>' -f $compareDate.ToString('dd/MM/yyyy'), $(
                    if ($exportToExcel.Count -eq 1) { 
                        'has been <b>1 delivery</b>' 
                    }
                    else { 
                        "have been <b>$($exportToExcel.Count) deliveries</b>"
                    }
                )
                $mailParams.Subject = '{0}, {1} {2}' -f  
                $s.Name, $exportToExcel.Count, $(
                    if ($exportToExcel.Count -eq 1) { 'delivery' }
                    else { 'deliveries' }
                )
                
                Send-MailAuthenticatedHC @mailParams
                #endregion
            }
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