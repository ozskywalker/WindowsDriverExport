# Windows Driver Export Script

- Ever needed to export 3rd party drivers to use in a Recovery image like Hiren's Boot CD?
- Ever needed a backup of that elusive 3rd party driver you can't source anymore?
- Or you've already exported drivers using DISM or Powershell but have no clue what mbedcomposite_x64.inf_amd64_67eb7e5c383aa11b & ssudobex.inf_amd64_1dd3a2846c69bed0 relates to?

This handy script exports your Windows drivers, then organizes them with meaningful folder names based on manufacturer, device class, and version information. It'll then give you a [summary text file at the end (sample here).](Sample_DriverMapping.txt)

This script has two modes:

- (default) **Export Mode**: Export all 3rd party drivers using `Export-WindowsDriver`
- **Parse-Only Mode**: Process an existing driver export folder without performing a new export

## Features
- **Folder Names you can understand**: Rename folders from obscure names like `mbtmdm.inf_amd64_72b3fac558336713` to `Schunid_Modem_mbtmdm_2.6.5.0`
- **Comprehensive Report**: Detailed text file with original & new folder name, manufacturer, device class (ie. Ports, Display, USB), Version and Class GUID
- **Collision Handling**: Automatically handle duplicate folder names

## Requirements

- Windows PowerShell 5.1 or PowerShell Core 6+
- Administrative privileges (required for driver operations)
- DISM PowerShell module (included with Windows)

## Usage

### Basic Export (Default)
```powershell
.\WindowsDriverExport.ps1
```
Exports drivers to `C:\DriversExport` and renames folders with meaningful names.

### Custom Export Location
```powershell
.\WindowsDriverExport.ps1 -ExportPath "D:\MyDriverBackup"
```

### Parse Existing Export (Skip Export)
```powershell
.\WindowsDriverExport.ps1 -ParseOnly -ExportPath "C:\ExistingDriverExport"
```

### Generate Report Only (No Renaming)
```powershell
.\WindowsDriverExport.ps1 -ParseOnly -NoRename -ExportPath "C:\ExistingDriverExport"
```
> [!TIP]
> Want to test this out? I've provided sample fake drivers in the [test_drivers](test_drivers/) & [real_test_drivers](real_test_drivers/) folder. Just clone this repo and try it out.

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `ExportPath` | String | `C:\DriversExport` | Target folder for export or existing export to parse |
| `ParseOnly` | Switch | False | Skip driver export and only process existing folder |
| `NoRename` | Switch | False | Generate report only, don't rename folders |
| `Verbose` | Switch | False | Enable verbose output |

## Output

### Folder Structure (After Renaming)

**Original DISM Export Structure:**
```
C:\DriversExport\
├── nvhda.inf_amd64_9fb9ca6ebbf0a797\
│   ├── nvhda.inf
│   └── [driver files...]
├── rt640x64.inf_amd64_neutral_f87f17b89c48c9e8\
│   ├── rt640x64.inf
│   └── [driver files...]
└── [other driver folders...]
```

**After Script Processing:**
```
C:\DriversExport\
├── NVIDIA_MEDIA_nvhda_1.3.40.21\
│   ├── nvhda.inf
│   └── [driver files...]
├── Realtek_Net_rt640x64_10.42.526.2020\
│   ├── rt640x64.inf
│   └── [driver files...]
└── DriverMapping.txt
```

### Driver Mapping Report
The script generates `DriverMapping.txt` containing:
- Original folder name → New folder name mapping
- Manufacturer, version, device class information
- Summary statistics by manufacturer and device class

#### Example Report Output

* [Full sample here](Sample_DriverMapping.txt)

```
Driver Export Mapping Report
Generated: 2024-01-15 14:30:22
Export Path: C:\DriversExport
Total Drivers Processed: 25

================================================================================

Original Folder: nvhda.inf_amd64_9fb9ca6ebbf0a797
New Folder:      NVIDIA_MEDIA_nvhda_1.3.40.21
INF File:        nvhda.inf
Manufacturer:    NVIDIA
Device Class:    MEDIA
Version:         1.3.40.21
Date:            01/15/2023
Class GUID:      {4d36e96c-e325-11ce-bfc1-08002be10318}
----------------------------------------

Original Folder: rt640x64.inf_amd64_neutral_f87f17b89c48c9e8
New Folder:      Realtek_Net_rt640x64_10.42.526.2020
INF File:        rt640x64.inf
Manufacturer:    Realtek Semiconductor Corp.
Device Class:    Network adapters
Version:         10.42.526.2020
Date:            05/15/2023
Class GUID:      {4D36E972-E325-11CE-BFC1-08002BE10318}
----------------------------------------

[... additional entries ...]

SUMMARY STATISTICS
========================================

Drivers by Manufacturer:
  NVIDIA: 8
  Intel: 6
  Realtek: 4
  AMD: 3
  Microsoft: 2
  Logitech: 2

Drivers by Device Class:
  Net: 8
  Display: 6
  System: 4
  HIDClass: 3
  Media: 2
  USB: 2
```

## Troubleshooting

### _"This script requires administrative privileges"_
Run PowerShell as Administrator.

### _"No driver folders found"_
Ensure the export path contains folders matching either:
- Modern DISM format: `infname.inf_arch_hash` (e.g., `nvhda.inf_amd64_9fb9ca6ebbf0a797`)
- Legacy format: `oemN.inf` (where N is a number)

### _Folder renaming failures_
Check file permissions and ensure no files in the folder are currently in use.

## License

[This script is licensed under MIT](LICENSE)