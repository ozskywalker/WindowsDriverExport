[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if ($_ -and -not (Test-Path -Path $_ -IsValid)) {
            throw "Invalid path format: $_"
        }
        return $true
    })]
    [string]$ExportPath = "C:\DriversExport",
    
    [Parameter(Mandatory=$false)]
    [switch]$ParseOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$NoRename
)

# Constants
$MAX_FOLDER_NAME_LENGTH = 50
$DRIVER_FOLDER_PATTERN = '^.+\.inf_(amd64|x86|arm64|neutral)_[a-f0-9]{16}$'
$OEM_FOLDER_PATTERN = '^oem\d+\.inf$'  # Legacy pattern for backwards compatibility
$INVALID_FOLDER_CHARS = '[<>:"/\\|?*]'

# Check admin privileges only when needed
if (-not $ParseOnly) {
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Error "Export mode requires administrative privileges. Use -ParseOnly to process existing exports without admin rights."
        exit 1
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message, 
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "ERROR" { Write-Error $logMessage }
        "WARNING" { Write-Warning $logMessage }
        default { Write-Host $logMessage }
    }
    
    if ($VerbosePreference -eq 'Continue') {
        Write-Verbose $logMessage
    }
}

function Export-Drivers {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Destination
    )
    
    Write-Log "Starting driver export to: $Destination"
    
    try {
        if (-not (Test-Path -Path $Destination)) {
            New-Item -ItemType Directory -Path $Destination -Force | Out-Null
            Write-Log "Created export directory: $Destination"
        }
        
        Write-Log "Exporting drivers using DISM..."
        $exportResult = Export-WindowsDriver -Online -Destination $Destination
        
        if ($exportResult -and $exportResult.Count -gt 0) {
            Write-Log "Driver export completed. Exported $($exportResult.Count) drivers."
            return $exportResult
        } else {
            Write-Log "No drivers were exported" "WARNING"
            return @()
        }
    }
    catch {
        Write-Log "Failed to export drivers: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Get-INFEncoding {
    param([string]$FilePath)
    
    try {
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
            return 'UTF8'
        }
        elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
            return 'Unicode'
        }
        elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
            return 'BigEndianUnicode'
        }
        else {
            return 'Default'  # ANSI/System default
        }
    }
    catch {
        return 'UTF8'  # Fallback
    }
}

function Parse-INFFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$INFFilePath
    )
    
    if (-not (Test-Path -Path $INFFilePath)) {
        Write-Log "INF file not found: $INFFilePath" "ERROR"
        return $null
    }
    
    try {
        # Detect encoding and read file
        $encoding = Get-INFEncoding -FilePath $INFFilePath
        $content = Get-Content -Path $INFFilePath -Encoding $encoding
        
        if (-not $content -or $content.Count -eq 0) {
            Write-Log "INF file is empty: $INFFilePath" "WARNING"
            return $null
        }
        
        # Initialize result with better defaults
        $fileName = Split-Path -Leaf $INFFilePath
        $folderName = Split-Path -Parent $INFFilePath | Split-Path -Leaf
        
        $driverInfo = [PSCustomObject]@{
            Provider = "Unknown"
            Version = "Unknown"
            Date = "Unknown"
            Class = "Unknown"
            ClassGuid = "Unknown"
            INFFile = $fileName
            FolderName = $folderName
            OriginalFolderName = ""
            NewFolderName = ""
            Renamed = $false
        }
        
        # Build string lookup table for references
        $stringTable = @{}
        foreach ($line in $content) {
            if ($line -match '^\s*([^=]+)\s*=\s*(.*)$') {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim() -replace '^["\s]*|["\s]*$', ''
                $stringTable[$key] = $value
            }
        }
        
        # Helper function to resolve string references
        function Resolve-StringReference {
            param([string]$Value, [hashtable]$StringTable)
            
            if ($Value -match '^%(.+)%$') {
                $key = $matches[1]
                if ($StringTable.ContainsKey($key)) {
                    return $StringTable[$key]
                }
                Write-Log "Unresolved string reference: %$key%" "WARNING"
                return $Value
            }
            return $Value
        }
        
        # Parse Provider
        $providerLine = $content | Where-Object { $_ -match '^\s*Provider\s*=' } | Select-Object -First 1
        if ($providerLine -and $providerLine -match '^\s*Provider\s*=\s*(.*)$') {
            $provider = $matches[1].Trim() -replace '^["\s]*|["\s]*$', ''
            $driverInfo.Provider = Resolve-StringReference -Value $provider -StringTable $stringTable
        }
        
        # Parse DriverVer
        $driverVerLine = $content | Where-Object { $_ -match '^\s*DriverVer\s*=' } | Select-Object -First 1
        if ($driverVerLine -and $driverVerLine -match '^\s*DriverVer\s*=\s*(.*)$') {
            $driverVer = $matches[1].Trim()
            if ($driverVer -match '^([^,]+),(.+)$') {
                $driverInfo.Date = $matches[1].Trim()
                $driverInfo.Version = $matches[2].Trim()
            } else {
                # DriverVer without comma (version only)
                $driverInfo.Version = $driverVer
            }
        }
        
        # Parse Class
        $classLine = $content | Where-Object { $_ -match '^\s*Class\s*=' } | Select-Object -First 1
        if ($classLine -and $classLine -match '^\s*Class\s*=\s*(.*)$') {
            $class = $matches[1].Trim() -replace '^["\s]*|["\s]*$', ''
            $driverInfo.Class = Resolve-StringReference -Value $class -StringTable $stringTable
        }
        
        # Parse ClassGUID with validation
        $classGuidLine = $content | Where-Object { $_ -match '^\s*ClassGuid\s*=' } | Select-Object -First 1
        if ($classGuidLine -and $classGuidLine -match '^\s*ClassGuid\s*=\s*(.*)$') {
            $guid = $matches[1].Trim() -replace '^["\s{]*|["\s}]*$', ''
            # Validate GUID format
            if ($guid -match '^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$') {
                $driverInfo.ClassGuid = $guid
            } else {
                Write-Log "Invalid GUID format in $INFFilePath`: $guid" "WARNING"
            }
        }
        
        return $driverInfo
    }
    catch {
        Write-Log "Error parsing INF file $INFFilePath`: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Get-SafeFolderName {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name
    )
    
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return "Unknown"
    }
    
    # Remove invalid characters and normalize
    $safeName = $Name -replace $INVALID_FOLDER_CHARS, '_'
    $safeName = $safeName -replace '\s+', '_'  # Replace multiple spaces with single underscore
    $safeName = $safeName.Trim('_')  # Remove leading/trailing underscores
    
    if ($safeName.Length -gt $MAX_FOLDER_NAME_LENGTH) {
        $safeName = $safeName.Substring(0, $MAX_FOLDER_NAME_LENGTH).TrimEnd('_')
    }
    
    return $safeName
}

function Generate-UniqueFolderName {
    param(
        [Parameter(Mandatory=$true)]
        [string]$BaseName,
        [Parameter(Mandatory=$true)]
        [string]$ParentPath,
        [Parameter(Mandatory=$true)]
        [hashtable]$UsedNames
    )
    
    $safeName = Get-SafeFolderName -Name $BaseName
    
    # If all components are unknown, use folder name as fallback
    if ($safeName -eq "Unknown_Unknown_Unknown") {
        $originalFolder = Split-Path -Leaf $ParentPath
        $safeName = "Unknown_$originalFolder"
    }
    
    $originalName = $safeName
    $counter = 1
    
    while ($UsedNames.ContainsKey($safeName.ToLower()) -or (Test-Path -Path (Join-Path $ParentPath $safeName))) {
        $safeName = "$originalName`_$counter"
        $counter++
        
        # Prevent infinite loops
        if ($counter -gt 9999) {
            $safeName = "$originalName`_$(Get-Random -Minimum 10000 -Maximum 99999)"
            break
        }
    }
    
    $UsedNames[$safeName.ToLower()] = $true
    return $safeName
}

function Process-DriverFolders {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ExportPath, 
        [Parameter(Mandatory=$true)]
        [bool]$RenameEnabled
    )
    
    Write-Log "Processing driver folders in: $ExportPath"
    
    if (-not (Test-Path -Path $ExportPath)) {
        throw "Export path does not exist: $ExportPath"
    }
    
    # Support both modern DISM format and legacy OEM format
    $driverFolders = Get-ChildItem -Path $ExportPath -Directory | Where-Object { 
        $_.Name -match $DRIVER_FOLDER_PATTERN -or $_.Name -match $OEM_FOLDER_PATTERN 
    }
    
    if ($driverFolders.Count -eq 0) {
        Write-Log "No driver folders found in $ExportPath (looking for .inf_arch_hash or oemN.inf patterns)" "WARNING"
        return @()
    }
    
    Write-Log "Found $($driverFolders.Count) driver folders to process"
    
    $processedDrivers = @()
    $usedNames = @{}
    $counter = 0
    $progressActivity = "Processing Drivers"
    
    try {
        foreach ($folder in $driverFolders) {
            $counter++
            $percentComplete = [math]::Round(($counter / $driverFolders.Count) * 100, 1)
            Write-Progress -Activity $progressActivity -Status "Processing folder $($folder.Name)" -PercentComplete $percentComplete
            
            $infFile = Get-ChildItem -Path $folder.FullName -Filter "*.inf" | Select-Object -First 1
            
            if (-not $infFile) {
                Write-Log "No INF file found in folder: $($folder.Name)" "WARNING"
                continue
            }
            
            $driverInfo = Parse-INFFile -INFFilePath $infFile.FullName
            
            if ($driverInfo) {
                $originalFolderName = $folder.Name
                $newFolderName = $originalFolderName
                
                if ($RenameEnabled) {
                    # Extract original INF name from folder if it's in modern format
                    $infName = ""
                    if ($originalFolderName -match '^(.+)\.inf_(amd64|x86|arm64|neutral)_[a-f0-9]{16}$') {
                        $infName = $matches[1]
                    }
                    
                    # Generate meaningful name with better fallbacks
                    if ($infName -and $infName -ne "oem") {
                        # Use original INF name if available and not generic OEM
                        $meaningfulName = "$($driverInfo.Provider)_$($driverInfo.Class)_$infName`_$($driverInfo.Version)"
                    } else {
                        $meaningfulName = "$($driverInfo.Provider)_$($driverInfo.Class)_$($driverInfo.Version)"
                    }
                    $newFolderName = Generate-UniqueFolderName -BaseName $meaningfulName -ParentPath $ExportPath -UsedNames $usedNames
                    
                    try {
                        Rename-Item -Path $folder.FullName -NewName $newFolderName -ErrorAction Stop
                        Write-Log "Renamed: $originalFolderName -> $newFolderName"
                        $driverInfo.Renamed = $true
                    } catch {
                        Write-Log "Failed to rename $originalFolderName to $newFolderName`: $($_.Exception.Message)" "ERROR"
                        $newFolderName = $originalFolderName
                        $driverInfo.Renamed = $false
                    }
                }
                
                $driverInfo.OriginalFolderName = $originalFolderName
                $driverInfo.NewFolderName = $newFolderName
                
                $processedDrivers += $driverInfo
            }
        }
    }
    finally {
        Write-Progress -Activity $progressActivity -Completed
    }
    
    return $processedDrivers
}

function Generate-Report {
    param(
        [Parameter(Mandatory=$true)]
        [array]$DriverData, 
        [Parameter(Mandatory=$true)]
        [string]$ExportPath
    )
    
    $reportPath = Join-Path $ExportPath "DriverMapping.txt"
    
    Write-Log "Generating driver mapping report: $reportPath"
    
    try {
        $report = @()
        $report += "Driver Export Mapping Report"
        $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $report += "Export Path: $ExportPath"
        $report += "Total Drivers Processed: $($DriverData.Count)"
        $report += ""
        $report += "=" * 80
        $report += ""
        
        foreach ($driver in $DriverData | Sort-Object Provider, Class) {
            $report += "Original Folder: $($driver.OriginalFolderName)"
            if ($driver.Renamed) {
                $report += "New Folder:      $($driver.NewFolderName)"
            }
            $report += "INF File:        $($driver.INFFile)"
            $report += "Manufacturer:    $($driver.Provider)"
            $report += "Device Class:    $($driver.Class)"
            $report += "Version:         $($driver.Version)"
            $report += "Date:            $($driver.Date)"
            $report += "Class GUID:      $($driver.ClassGuid)"
            $report += "-" * 40
            $report += ""
        }
        
        # Add summary statistics
        $report += ""
        $report += "SUMMARY STATISTICS"
        $report += "=" * 40
        
        if ($DriverData.Count -gt 0) {
            $manufacturerStats = $DriverData | Group-Object Provider | Sort-Object Count -Descending
            $report += ""
            $report += "Drivers by Manufacturer:"
            foreach ($stat in $manufacturerStats) {
                $report += "  $($stat.Name): $($stat.Count)"
            }
            
            $classStats = $DriverData | Group-Object Class | Sort-Object Count -Descending
            $report += ""
            $report += "Drivers by Device Class:"
            foreach ($stat in $classStats) {
                $report += "  $($stat.Name): $($stat.Count)"
            }
        }
        
        $report | Out-File -FilePath $reportPath -Encoding UTF8 -Force
        Write-Log "Report generated successfully: $reportPath"
    }
    catch {
        Write-Log "Failed to generate report: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Main execution
$ErrorActionPreference = 'Stop'

try {
    Write-Log "Driver Export and Rename Script Started"
    Write-Log "Parameters - ExportPath: $ExportPath, ParseOnly: $ParseOnly, NoRename: $NoRename"
    
    # Validate export path
    if ($ParseOnly) {
        if (-not (Test-Path -Path $ExportPath)) {
            throw "Export path does not exist: $ExportPath"
        }
        Write-Log "Parse-only mode: Processing existing export at $ExportPath"
    } else {
        Write-Log "Export mode: Will export drivers to $ExportPath"
        $null = Export-Drivers -Destination $ExportPath
    }
    
    # Process driver folders
    $renameEnabled = -not $NoRename
    $processedDrivers = Process-DriverFolders -ExportPath $ExportPath -RenameEnabled $renameEnabled
    
    if ($processedDrivers.Count -eq 0) {
        Write-Log "No drivers were processed successfully" "WARNING"
    } else {
        Write-Log "Successfully processed $($processedDrivers.Count) drivers"
        
        # Generate report
        Generate-Report -DriverData $processedDrivers -ExportPath $ExportPath
        
        if ($renameEnabled) {
            $renamedCount = ($processedDrivers | Where-Object { $_.Renamed }).Count
            Write-Log "Renamed $renamedCount folders with meaningful names"
        }
    }
    
    Write-Log "Script completed successfully"
    exit 0
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" "ERROR"
    exit 1
}