#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
    }
    
    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile', 'ScriptName') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
            ($Subject -eq 'FAILURE')
        }    
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'
    
            .$testScript @testNewParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It 'MailTo is missing' {
                @{
                    Suppliers = @(
                        @{
                            Name   = 'Picard'
                            Path   = 'TestDrive:/'
                            MailTo = 'bob@contoso.com'
                        }
                    )
                } | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'MailTo' addresses found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Suppliers is missing' {
                @{
                    MailTo = @('bob@contoso.com')
                } | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Suppliers' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'Property Suppliers' {
                It 'Path is missing' {
                    @{
                        MailTo    = @('bob@contoso.com')
                        Suppliers = @(
                            @{
                                Name   = 'Picard'
                                # Path   = 'TestDrive:/'
                                MailTo = 'bob@contoso.com'
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Path' is missing in 'Suppliers'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Path does not exist' {
                    @{
                        MailTo    = @('bob@contoso.com')
                        Suppliers = @(
                            @{
                                Name   = 'Picard'
                                Path   = 'C:/notExisting'
                                MailTo = 'bob@contoso.com'
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*'Path' folder 'C:/notExisting' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Name is missing' {
                    @{
                        MailTo    = @('bob@contoso.com')
                        Suppliers = @(
                            @{
                                # Name   = 'Picard'
                                Path   = 'TestDrive:/'
                                MailTo = 'bob@contoso.com'
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Name' is missing in 'Suppliers'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'MailTo is missing' {
                    @{
                        MailTo    = @('bob@contoso.com')
                        Suppliers = @(
                            @{
                                Name = 'Picard'
                                Path = 'TestDrive:/'
                                # MailTo = 'bob@contoso.com'
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'MailTo' is missing in 'Suppliers'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'MailTo is not an email address' {
                    @{
                        MailTo    = @('bob@contoso.com')
                        Suppliers = @(
                            @{
                                Name   = 'Picard'
                                Path   = 'TestDrive:/'
                                MailTo = 'invalid'
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*'MailTo' value 'invalid' is not a valid e-mail address for supplier 'Picard'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
        }
    }
}
Describe 'when all tests pass' {
    BeforeAll {
        $testAscFile = @"
BE1021058802552104737363                    0022016630Faber W Krommenie                  Rosariumlaan 47                    KROMMENI                           000000000000103464CEM I 42,5 N BULK                       29.700202203142022031507150092BJT9              CNLSS128
NL1121058805192104737268                    0021700679MEBIN Tessel DENBOSCH              Tesselschadestraat 30              's-Hertogenbosch                   000000000000103415CEM III/B 42,5 N LH NCR BULK            37.7802022031520220315060000DUMSIMONS01         C
"@

        $testExportedExcelRows = @(
            @{
                Plant               = 'BE10'
                ShipmentNumber      = 2105880255
                DeliveryNumber      = 2104737363
                ShipToNumber        = 22016630
                ShipToName          = 'Faber W Krommenie'
                Address             = 'Rosariumlaan 47'
                City                = 'KROMMENI'
                MaterialNumber      = 103464
                MaterialDescription = 'CEM I 42,5 N BULK'
                Tonnage             = 29.700
                LoadingDate         = Get-Date('3/14/2022')
                DeliveryDate        = Get-Date('3/15/2022 7:15:00 AM')
                TruckID             = '92BJT9'
                PickingStatus       = 'C'
                SiloBulkID          = 'NLSS128'
            }
            @{
                Plant               = 'NL11'
                ShipmentNumber      = 2105880519
                DeliveryNumber      = 2104737268
                ShipToNumber        = 21700679
                ShipToName          = 'MEBIN Tessel DENBOSCH'
                Address             = 'Tesselschadestraat 30'
                City                = "'s-Hertogenbosch"
                MaterialNumber      = 103415
                MaterialDescription = 'CEM III/B 42,5 N LH NCR BULK'
                Tonnage             = 37.780
                LoadingDate         = Get-Date('3/15/2022')
                DeliveryDate        = Get-Date('3/15/2022 6:00:00 AM')
                TruckID             = 'DUMSIMONS01'
                PickingStatus       = 'C'
                SiloBulkID          = ''
            }
        )

        $testAscFileOutParams = @{
            FilePath = (New-Item "TestDrive:/Test.asc" -ItemType File).FullName
            Encoding = 'utf8'
        }
        $testAscFile | Out-File @testAscFileOutParams

        @{
            MailTo    = 'bob@contoso.com'
            Suppliers = @(
                @{
                    Name   = 'Picard'
                    Path   = 'TestDrive:/'
                    MailTo = 'bob@contoso.com'
                }
            )
        } | ConvertTo-Json | Out-File @testOutParams
        
        .$testScript @testParams
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Data'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.ShipmentNumber -eq $testRow.ShipmentNumber
                }
                @(
                    'Plant', 'DeliveryNumber', 'ShipToNumber', 'ShipToName',
                    'Address', 'City', 'MaterialNumber', 'MaterialDescription',
                    'Tonnage', 'LoadingDate', 'TruckID', 'PickingStatus', 
                    'SiloBulkID'
                ) | ForEach-Object {
                    $actualRow.$_ | Should -Be $testRow.$_
                }
            }
        }
    }
} -Tag test