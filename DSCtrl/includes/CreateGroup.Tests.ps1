$here = (Split-Path -Parent $MyInvocation.MyCommand.Path)
. $here\CreateGroup.ps1

Function Invoke-TestProxy {}

Describe 'ConvertFrom-GroupName' {

    Context "Handling invalid input" {
        It 'Given an empty string, it returns null' {
            $shouldBeNull = ConvertFrom-GroupName -Name ' '
            $shouldBeNull | Should Be $null
        }
    
        It "Given invalid parameter -Name 'I like Star Wars!', it returns `$null" {
            $shouldBeNull = ConvertFrom-GroupName -Name 'I like Star Wars!'
            $shouldBeNull | Should Be $null
        }
    }
    
    Context "Correctly identifies group types" {
        It "Given valid -Name '<Name>', it returns type as '<Expected>'" -TestCases @(
            @{ Name = 'M.UTOSPA.UTO.Groups.CMP'; Expected = 'Computer' }
            @{ Name = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.DS'  ; Expected = 'Computer' }
            @{ Name = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.UCC.Poly'  ; Expected = 'Computer' }
            @{ Name = 'M.UTOSPA.UTO.Groups.PRT.ITCCS.DS.ASUPRINT1_USI-TS-MCL117'   ; Expected = 'Printer' }
            @{ Name = 'M.UTOSPA.UTO.Groups.PRT.ITCSS.P.POLYPRINT1_EPRT506C'; Expected = 'Printer' }
            @{ Name = 'M.UTOSPA.UTO.Groups.SHR-RW.itfs1_uto_groups_support_ats'; Expected = 'Share' }
            @{ Name = 'M.UTOSPA.UTO.Groups.USR.ITCSS.DS.POLY.STAFF'; Expected = 'User' }
            @{ Name = 'M.UTOSPA.UTO.Groups.GPO.DS_CHROME_EXTENSION'; Expected = 'GPO' }
        ) {
            param ($Name, $Expected)
            $parseResult = ConvertFrom-GroupName -Name $Name
            $parseResult.type | Should Be $Expected
        }
    }

    Context "Correctly identifies parent group names" {
        It "Given valid -Name '<Name>', it returns type as '<Expected>'" -TestCases @(
            @{ Name = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.DS'  ; Expected = 'M.UTOSPA.UTO.Groups.CMP.ITCSS' }
            @{ Name = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.UCC.Poly'  ; Expected = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.UCC' }
            @{ Name = 'M.UTOSPA.UTO.Groups.PRT.ITCCS.DS.ASUPRINT1_USI-TS-MCL117'   ; Expected = 'M.UTOSPA.UTO.Groups.CMP.ITCCS.DS' }
            @{ Name = 'M.UTOSPA.UTO.Groups.SHR-RW.itfs1_uto_groups_support_ats'; Expected = '' }
            @{ Name = 'M.UTOSPA.UTO.Groups.USR.ITCSS.DS.POLY.STAFF'; Expected = 'M.UTOSPA.UTO.Groups.USR.ITCSS.DS.POLY' }
            @{ Name = 'M.UTOSPA.UTO.Groups.GPO.DS_CHROME_EXTENSION'; Expected = '' }
            @{ Name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_UTO-USB2605-C3525i'; Expected = 'M.UTOSPA.UTO.Groups.CMP' }
            @{ Name = 'M.UTOSPA.UTO.Groups.PRT.ITCSS.DS.DPCPRINT1_DPCIT_L162_DellMFP'; Expected = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.DS' }
        ) {
            param ($Name, $Expected)
            $parseResult = ConvertFrom-GroupName -Name $Name
            $parseResult.parent | Should Be $Expected
        }
    }

    Context "Correctly identifies Unit names" {
        It "Given valid -Name '<Name>', it returns unit name as '<Expected>'" -TestCases @(
            @{ Name = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.UCC.Poly'  ; Expected = 'UTO' }
            @{ Name = 'M.UTOSPA.Alumni.Groups.PRT.asuprint1_printer'   ; Expected = 'Alumni' }
        ) {
            param ($Name, $Expected)
            $parseResult = ConvertFrom-GroupName -Name $Name
            $parseResult.unit | Should Be $Expected
        }
    }

    Context "Correctly identifies Resource paths" {
        It "Given valid -Name '<Name>', it returns Resource path as '<Expected>'" -TestCases @(
            @{
                Name = 'M.UTOSPA.UTO.Groups.PRT.POLYPRINT1_EPRT177'
                Expected = '\\POLYPRINT1\EPRT177'
            }
            @{ 
                Name = 'M.UTOSPA.UTO.Groups.PRT.POLYPRINT1_EPRT177_super_printer_9000'
                Expected = '\\POLYPRINT1\EPRT177_super_printer_9000'
            }
            @{ 
                Name = 'M.UTOSPA.UTO.Groups.PRT.ITCSS.DS.DPCPRINT1_DCSS_MercC_117_Canon'
                Expected = '\\DPCPRINT1\DCSS_MercC_117_Canon'
            }
            @{ 
                Name = 'M.UTOSPA.UTO.Groups.SHR-RW.ITCSS.DS.POLY.ITFS1_POLYUTO$_UTO_DS'
                Expected = '\\ITFS1\POLYUTO$_UTO_DS'
            }
            @{ 
                Name = 'M.UTOSPA.UTO.Groups.SHR-DA.itfs1_uto_groups_support_dstempe~1'
                Expected = '\\itfs1\uto_groups_support_dstempe'
            }
        ) {
            param ($Name, $Expected)
            $parseResult = ConvertFrom-GroupName -Name $Name
            $parseResult.resource | Should Be $Expected
        }
    }
}

Describe 'ConvertTo-GroupName' {

    Context "Creating computer groups" {
        It "Able to generate '<Expected>'" -TestCases @(
            @{
                params = @{
                    GroupType = 'Computer'
                    GroupUnit = 'UTO'
                    GroupParent = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.DS'
                    GroupSuffix = 'DSO'
                }
                Expected = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.DS.DSO'
            }
            @{
                params = @{
                    GroupType = 'Computer'
                    GroupUnit = 'UTO'
                    GroupSuffix = 'ITCSS'
                }
                Expected = 'M.UTOSPA.UTO.Groups.CMP.ITCSS'
            }
            @{
                params = @{
                    GroupType = 'Computer'
                    GroupUnit = 'UTO'
                    GroupParent = 'M.UTOSPA.UTO.Groups.CMP.ITCSS'
                    GroupSuffix = 'EMAIL'
                }
                Expected = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.EMAIL'
            }
            @{
                params = @{
                    GroupType = 'Computer'
                    GroupUnit = 'UTO'
                    GroupParent = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.DS'
                    GroupSuffix = 'CLAS'
                }
                Expected = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.DS.CLAS'
            }
        ) {
            param ($params, $Expected)
            $parseResult = ConvertTo-GroupName @params
            $parseResult | Should Be $Expected
        }
    }

    Context "Creating User groups" {
        It "Able to generate '<Expected>'" -TestCases @(
            @{
                params = @{
                    GroupType = 'User'
                    GroupUnit = 'UTO'
                    GroupParent = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.DS.POLY'
                    GroupSuffix = 'Staff'
                }
                Expected = 'M.UTOSPA.UTO.Groups.USR.ITCSS.DS.POLY.STAFF'
            }
            @{
                params = @{
                    GroupType = 'User'
                    GroupUnit = 'UTO'
                    GroupSuffix = 'ITCSS'
                }
                Expected = 'M.UTOSPA.UTO.Groups.USR.ITCSS'
            }
            @{
                params = @{
                    GroupType = 'User'
                    GroupUnit = 'UTO'
                    GroupParent = 'M.UTOSPA.UTO.Groups.CMP.ITCSS'
                    GroupSuffix = 'Email'
                }
                Expected = 'M.UTOSPA.UTO.Groups.USR.ITCSS.Email'
            }
            @{
                params = @{
                    GroupType = 'User'
                    GroupUnit = 'UTO'
                    GroupParent = 'M.UTOSPA.UTO.Groups.CMP.ITCSS'
                    GroupSuffix = 'Example'
                    CampusInitial = 'P'
                }
                Expected = 'M.UTOSPA.UTO.Groups.USR.ITCSS.Example.P'
            }
        ) {
            param ($params, $Expected)
            $parseResult = ConvertTo-GroupName @params
            $parseResult | Should Be $Expected
        }
    }

    Context "Creating Printer groups" {
        Mock -CommandName 'Invoke-TestProxy' -MockWith {
            return (
                [PSCustomObject]@{
                    name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_USI-TS-MCL134-LJ2100-9000SUP~1'
                },
                [PSCustomObject]@{
                    name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_USI-TS-MCL134-LJ2100-9000SUP~2'
                },
                [PSCustomObject]@{
                    name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_USI-TS-MCL134-LJ2100-9000SUP~3'
                },
                [PSCustomObject]@{
                    name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_IDont_Always_Name_My_Printer~1'
                },
                [PSCustomObject]@{
                    name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_ABCDEFGHIJKLMNOPQRSTUVWXYZ00~1'
                },
                [PSCustomObject]@{
                    name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_ABCDEFGHIJKLMNOPQRSTUVWXYZ00~2'
                },
                [PSCustomObject]@{
                    name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_ABCDEFGHIJKLMNOPQRSTUVWXYZ00~3'
                },
                [PSCustomObject]@{
                    name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_ABCDEFGHIJKLMNOPQRSTUVWXYZ00~4'
                },
                [PSCustomObject]@{
                    name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_ABCDEFGHIJKLMNOPQRSTUVWXYZ00~5'
                },
                [PSCustomObject]@{
                    name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_ABCDEFGHIJKLMNOPQRSTUVWXYZ00~6'
                },
                [PSCustomObject]@{
                    name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_ABCDEFGHIJKLMNOPQRSTUVWXYZ00~7'
                },
                [PSCustomObject]@{
                    name = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_ABCDEFGHIJKLMNOPQRSTUVWXYZ00~8'
                }
            )
        }
        It "Able to generate '<Expected>'" -TestCases @(
            @{
                params = @{
                    GroupType = 'Printer'
                    GroupUnit = 'UTO'
                    ResourceServer = 'ASUPRINT1'
                    ResourcePath = 'UTO-USB2370-C3525i'
                    SkipADQuery = $true
                }
                Expected = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_UTO-USB2370-C3525i'
            }
            @{
                params = @{
                    GroupType = 'Printer'
                    GroupUnit = 'UTO'
                    GroupParent = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.DS'
                    ResourceServer = 'ASUPRINT1'
                    ResourcePath = 'USI-TS-MCL106-LJ8100'
                    SkipADQuery = $true
                }
                Expected = 'M.UTOSPA.UTO.Groups.PRT.ITCSS.DS.ASUPRINT1_USI-TS-MCL106-LJ8100'
            }
            @{
                params = @{
                    GroupType = 'Printer'
                    GroupUnit = 'UTO'
                    GroupParent = 'M.UTOSPA.UTO.Groups.CMP.ITCSS'
                    ResourceServer = 'POLYPRINT1'
                    ResourcePath = 'EPRT171'
                    CampusInitial = 'P'
                    SkipADQuery = $true
                }
                Expected = 'M.UTOSPA.UTO.Groups.PRT.ITCSS.P.POLYPRINT1_EPRT171'
            }
            @{
                params = @{
                    GroupType = 'Printer'
                    GroupUnit = 'UTO'
                    ResourceServer = 'POLYPRINT1'
                    ResourcePath = 'A_Printer_With_A_Really_Long_Name'
                    SkipADQuery = $true
                }
                Expected = 'M.UTOSPA.UTO.Groups.PRT.POLYPRINT1_A_Printer_With_A_Really_Lon~1'
            }
            @{
                params = @{
                    GroupType = 'Printer'
                    GroupUnit = 'UTO'
                    ResourceServer = 'ASUPRINT1'
                    ResourcePath = 'USI-TS-MCL134-LJ2100-9000SUPER-PRINTER'
                }
                Expected = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_USI-TS-MCL134-LJ2100-9000SUP~4'
            }
            @{
                params = @{
                    GroupType = 'Printer'
                    GroupUnit = 'UTO'
                    ResourceServer = 'ASUPRINT1'
                    ResourcePath = 'IDont_Always_Name_My_Printer_but_when_I_do'
                }
                Expected = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_IDont_Always_Name_My_Printer~2'
            }
            @{
                params = @{
                    GroupType = 'Printer'
                    GroupUnit = 'UTO'
                    ResourceServer = 'ASUPRINT1'
                    ResourcePath = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0000'
                }
                Expected = 'M.UTOSPA.UTO.Groups.PRT.ASUPRINT1_ABCDEFGHIJKLMNOPQRSTUVWXYZ00~9'
            }
        ) {
            param ($params, $Expected)
            $parseResult = ConvertTo-GroupName @params
            $parseResult | Should Be $Expected
        }
    }

    Context "Creating Share groups" {
        It "Able to generate '<Expected>'" -TestCases @(
            @{
                params = @{
                    GroupType = 'Share'
                    GroupUnit = 'UTO'
                    ResourceServer = 'ITFS1'
                    ResourcePath = 'path$\to\some\share'
                    AccessType = 'RW'
                }
                Expected = 'M.UTOSPA.UTO.Groups.SHR-RW.ITFS1_path$_to_some_share'
            }
            @{
                params = @{
                    GroupType = 'Share'
                    GroupUnit = 'UTO'
                    GroupParent = 'M.UTOSPA.UTO.Groups.CMP.ITCSS.DS.POLY'
                    ResourceServer = 'ITFS1'
                    ResourcePath = 'POLYUTO$\UTO_DS'
                    AccessType = 'DA'
                }
                Expected = 'M.UTOSPA.UTO.Groups.SHR-DA.ITCSS.DS.POLY.ITFS1_POLYUTO$_UTO_DS'
            }
            @{
                params = @{
                    GroupType = 'Share'
                    GroupUnit = 'UTO'
                    ResourceServer = 'ITFS1'
                    ResourcePath = 'POLYUTO$\UTO_DS [chadd@poly]'
                    AccessType = 'RW'
                }
                Expected = 'M.UTOSPA.UTO.Groups.SHR-RW.ITFS1_POLYUTO$_UTO_DS chaddpoly'
            }
        ) {
            param ($params, $Expected)
            $parseResult = ConvertTo-GroupName @params
            $parseResult | Should Be $Expected
        }
    }

    Context "Creating GPO groups" {
        It "Able to generate '<Expected>'" -TestCases @(
            @{
                params = @{
                    GroupType = 'GPO'
                    GroupUnit = 'UTO'
                    ResourcePath = 'DS_CHROME_EXTENSION'
                    AccessType = 'RW'
                }
                Expected = 'M.UTOSPA.UTO.Groups.GPO.DS_CHROME_EXTENSION'
            }
        ) {
            param ($params, $Expected)
            $parseResult = ConvertTo-GroupName @params
            $parseResult | Should Be $Expected
        }
    }
}