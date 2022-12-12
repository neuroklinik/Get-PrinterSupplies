function Get-PrinterSupplies {
    param (
        [Parameter(Mandatory)][String]$hostname,
        [Parameter()][String]$snmpreadcommunity = 'public'
    )
    $snmp = New-Object -ComObject oleprn.olesnmp
    $snmp.open($hostname, $snmpreadcommunity, 5, 6000)
    $printermodel = $snmp.get('.1.3.6.1.2.1.25.3.2.1.3.1')
    $suppliesClasses = $snmp.GetTree('.1.3.6.1.2.1.43.11.1.1.4')
    $suppliesTypes = $snmp.GetTree('.1.3.6.1.2.1.43.11.1.1.5')
    $suppliesDescriptions = $snmp.GetTree('.1.3.6.1.2.1.43.11.1.1.6')
    $suppliesUnits = $snmp.GetTree('.1.3.6.1.2.1.43.11.1.1.7')
    $suppliesMaxCapacity = $snmp.GetTree('.1.3.6.1.2.1.43.11.1.1.8')
    $suppliesCurrentLevel = $snmp.GetTree('.1.3.6.1.2.1.43.11.1.1.9')
    $supplydata = @()
    for ($i = 0; $i -lt $suppliesclasses.length / 2; $i++) {
        $supplydata += [pscustomobject]@{
            "Hostname"          = $hostname;
            "Printer Model"     = $printermodel;
            "Description"       = $suppliesDescriptions[1, $i];
            "Class"             = switch ($suppliesClasses[1, $i]) {
                1 { "Other" }
                3 { "Consumed" }
                4 { "Filled" }
            };
            "Type"              = switch ($suppliesTypes[1, $i]) {
                1 { "Other" }
                2 { "Unknown" }
                3 { "Toner" }
                4 { "Waste Toner" }
                5 { "Ink" }
                6 { "Ink Cartridge" }
                7 { "Ink Ribbon" }
                8 { "Waste Ink" }
                9 { "OPC Drum" }
                10 { "Developer" }
                11 { "Fuser Oil" }
                12 { "Solid Wax" }
                13 { "Ribbon Wax" }
                14 { "Waste Wax" }
                15 { "Fuser" }
                16 { "Corona Wire" }
                17 { "Fuser Oil Wick" }
                18 { "Cleaner Unit" }
                19 { "Fuser Cleaning Pad" }
                20 { "Transfer Unit" }
                21 { "Toner Cartridge" }
                22 { "Fuser Oiler" }
                23 { "Water" }
                24 { "Waste Water" }
                25 { "Glue Water Additive" }
                26 { "Waste Paper" }
                27 { "Binding Supply" }
                28 { "Banding Supply" }
                29 { "Stitching Wire" }
                30 { "Shrink Wrap" }
                31 { "Paper Wrap" }
                32 { "Staples" }
                33 { "Inserts" }
                34 { "Covers" }
            };
            "Units"             = switch ($suppliesUnits[1, $i]) {
                1 { "Other" }
                2 { "Unknown" }
                3 { "Ten Thousandths of Inches" }
                4 { "Micrometers" }
                7 { "Impressions" }
                8 { "Sheets" }
                11 { "Hours" }
                12 { "Thousandths of Ounces" }
                13 { "Tenths of Grams" }
                14 { "Hundredths of Fluid Ounces" }
                15 { "Tenths of Milliliters" }
                16 { "Feet" }
                17 { "Meters" }
                18 { "Items" }
                19 { "Percent" }
            };
            "Max Capacity"      = switch ($suppliesMaxCapacity[1, $i]) {
                -1 { "Other" }
                -2 { "Unknown" }
                Default { $suppliesMaxCapacity[1, $i] }
            };
            "Current Level"     = switch ($suppliesCurrentLevel[1, $i]) {
                -1 { "Other" }
                -2 { "Unknown" }
                -3 { "Undepleted" }
                Default { $suppliesCurrentLevel[1, $i] }
            };
#            "Percent Remaining" = Try { (($suppliesMaxCapacity[1, $i] -lt 0) -or ($suppliesCurrentLevel[1, $i] -lt 0))?"N/A":[int]($suppliesCurrentLevel[1, $i] / $suppliesMaxCapacity[1, $i] * 100) } catch [System.DivideByZeroException] { "Div0Error" };
            "Percent Remaining" = Try { if (($suppliesMaxCapacity[1, $i] -lt 0) -or ($suppliesCurrentLevel[1, $i] -lt 0)) {"N/A"} else {[int]($suppliesCurrentLevel[1, $i] / $suppliesMaxCapacity[1, $i] * 100)} } catch [System.DivideByZeroException] { "Div0Error" };
        }
    }
    $snmp.close()
    Return $supplydata
}