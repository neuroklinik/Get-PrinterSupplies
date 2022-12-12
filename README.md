# Get-PrinterSupplies
A PowerShell function to collect consumables data from a printer via SNMP.
## Parameters
* -hostname (mandatory): the hostname or IP address of the printer as a string
* -snmpreadcommunity (optional): the SNMP read community string for the printer, uses *public* by default
## Operation
Contacts the target printer over SNMP, collects and interprets consumables data, and outputs an array of custom PowerShell objects with properties and values of the consumables.

Collects:
* Printer Model
* Supplies Descriptions, Classes, Types, Measurement Units, Maximum Capacity and Current Level

Adds:
* A calculated percent remaining.
