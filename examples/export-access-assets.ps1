[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [ValidateSet('Full', 'StartupProbe')]
    [string]$Mode = 'Full',

    [ValidateRange(10, 900)]
    [int]$OpenTimeoutSeconds = 90,

    [SecureString]$DatabasePassword,

    [switch]$IncludeLinkedTableDetails,

    [switch]$KeepWorkingCopy,

    [ValidateSet('Fail', 'RestrictedLocal')]
    [string]$SensitiveOutputPolicy = 'Fail',

    [switch]$AcknowledgeRestrictedOutput
)

$ErrorActionPreference = 'Stop'

if (-not ('AccessExportNative' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class AccessExportNative
{
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int virtualKey);

    public const byte VK_SHIFT = 0x10;
    public const uint KEYEVENTF_KEYUP = 0x0002;

    public static void ShiftDown()
    {
        keybd_event(VK_SHIFT, 0, 0, UIntPtr.Zero);
    }

    public static void ShiftUp()
    {
        keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
'@
}

try {
    [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
}
catch {
    # Windows PowerShell already exposes the Windows code pages.
}

$script:LogLines = New-Object System.Collections.Generic.List[string]
$script:Errors = New-Object System.Collections.Generic.List[object]
$script:Manifest = New-Object System.Collections.Generic.List[object]
$script:Discovered = [ordered]@{
    forms = 0
    reports = 0
    modules = 0
    macros = 0
    queries = 0
    data_macros = 0
}
$script:Exported = [ordered]@{
    forms = 0
    reports = 0
    modules = 0
    macros = 0
    queries = 0
    data_macros = 0
}
$script:SkippedByMode = New-Object System.Collections.Generic.List[string]
$script:DataMacroTablesProbed = 0
$script:DataMacroAbsentCount = 0
$script:DataMacroProbeErrorCount = 0
$script:SensitiveTokenHits = New-Object System.Collections.Generic.List[object]
$script:EmptyObjects = New-Object System.Collections.Generic.List[object]

function Add-LogLine {
    param([Parameter(Mandatory = $true)][string]$Message)

    $timestamp = (Get-Date).ToString('o')
    [void]$script:LogLines.Add("$timestamp $Message")
}

function Protect-Message {
    param([AllowEmptyString()][string]$Message)

    if ($null -eq $Message) {
        return ''
    }

    $safe = $Message
    if ($script:SourcePath) {
        $safe = [regex]::Replace($safe, [regex]::Escape($script:SourcePath), '<source-db>', 'IgnoreCase')
        $safe = [regex]::Replace($safe, [regex]::Escape((Split-Path -Parent $script:SourcePath)), '<source-dir>', 'IgnoreCase')
    }
    if ($script:OutputPath) {
        $safe = [regex]::Replace($safe, [regex]::Escape($script:OutputPath), '<output-dir>', 'IgnoreCase')
    }
    $userProfile = [Environment]::GetEnvironmentVariable('USERPROFILE', 'Process')
    if (-not [string]::IsNullOrWhiteSpace($userProfile)) {
        $safe = [regex]::Replace($safe, [regex]::Escape($userProfile), '<user-profile>', 'IgnoreCase')
    }

    $safe = [regex]::Replace(
        $safe,
        '(?i)(PWD|PASSWORD|UID|USER ID|ACCESS TOKEN)\s*=\s*[^;\s]+',
        '$1=<redacted>'
    )
    return $safe
}

function Get-EnvironmentFingerprint {
    $variableNames = @(
        'PATHEXT', 'COMSPEC', 'SystemRoot', 'ProgramFiles', 'ProgramFiles(x86)',
        'CommonProgramFiles', 'CommonProgramFiles(x86)', 'LOCALAPPDATA',
        'ProgramData', 'TEMP', 'USERPROFILE', 'APPDATA'
    )

    $variables = [ordered]@{}
    foreach ($name in $variableNames) {
        $value = [Environment]::GetEnvironmentVariable($name, 'Process')
        $variables[$name] = [pscustomobject]@{
            present = -not [string]::IsNullOrWhiteSpace($value)
            value = Protect-Message -Message $value
        }
    }

    return [pscustomobject]@{
        process_environment_variable_count = [Environment]::GetEnvironmentVariables('Process').Count
        powershell_version = $PSVersionTable.PSVersion.ToString()
        process_bitness = if ([Environment]::Is64BitProcess) { '64-bit' } else { '32-bit' }
        operating_system_bitness = if ([Environment]::Is64BitOperatingSystem) { '64-bit' } else { '32-bit' }
        variables = $variables
    }
}

function Add-ExportError {
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [string]$ObjectType = '',
        [string]$ObjectName = '',
        [Parameter(Mandatory = $true)][string]$Message
    )

    $safeMessage = Protect-Message -Message $Message
    [void]$script:Errors.Add([pscustomobject]@{
        stage = $Stage
        object_type = $ObjectType
        object_name = $ObjectName
        message = $safeMessage
    })
    Add-LogLine -Message "ERROR [$Stage] $ObjectType $ObjectName $safeMessage"
}

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
        catch {
            # Cleanup is best effort; the exact Access PID is checked later.
        }
    }
}

function Test-ExceptionErrorNumber {
    param(
        [Parameter(Mandatory = $true)][Exception]$Exception,
        [Parameter(Mandatory = $true)][int[]]$ExpectedNumbers
    )

    $current = $Exception
    while ($null -ne $current) {
        if (($current.HResult -band 0xFFFF) -in $ExpectedNumbers) {
            return $true
        }
        $current = $current.InnerException
    }
    return $false
}

function Write-Utf8Text {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [AllowEmptyString()][string]$Text
    )

    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, $Text, $utf8)
}

function Write-Utf8Json {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object]$Value
    )

    $json = ConvertTo-Json -InputObject $Value -Depth 30
    Write-Utf8Text -Path $Path -Text ($json + [Environment]::NewLine)
}

function Get-SafeFileName {
    param(
        [Parameter(Mandatory = $true)][int]$Index,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Extension
    )

    $safeName = $Name
    foreach ($character in [IO.Path]::GetInvalidFileNameChars()) {
        $safeName = $safeName.Replace([string]$character, '_')
    }
    $safeName = $safeName.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = 'unnamed'
    }
    if ($safeName.Length -gt 100) {
        $safeName = $safeName.Substring(0, 100)
    }

    return ('{0:D4}_{1}{2}' -f $Index, $safeName, $Extension)
}

function Get-RelativePathForManifest {
    param([Parameter(Mandatory = $true)][string]$Path)

    $root = $script:OutputPath.TrimEnd('\') + '\'
    if (-not $Path.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        throw "File is outside the output directory: $Path"
    }
    return $Path.Substring($root.Length).Replace('\', '/')
}

function Add-ManifestFile {
    param(
        [Parameter(Mandatory = $true)][string]$Tier,
        [Parameter(Mandatory = $true)][string]$SourceKind,
        [Parameter(Mandatory = $true)][string]$Category,
        [string]$ObjectType = '',
        [string]$ObjectName = '',
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Encoding,
        [switch]$AllowEmpty
    )

    $file = Get-Item -LiteralPath $Path
    if ($file.Length -eq 0 -and -not $AllowEmpty) {
        throw "Exported file is empty: $Path"
    }

    [void]$script:Manifest.Add([pscustomobject]@{
        tier = $Tier
        source_kind = $SourceKind
        category = $Category
        object_type = $ObjectType
        object_name = $ObjectName
        relative_path = Get-RelativePathForManifest -Path $Path
        byte_length = $file.Length
        content_state = if ($file.Length -eq 0) { 'empty' } else { 'nonempty' }
        encoding = $Encoding
        sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    })
}

function Convert-ToReviewUtf8 {
    param(
        [Parameter(Mandatory = $true)][string]$NativePath,
        [Parameter(Mandatory = $true)][string]$ReviewPath,
        [switch]$AllowEmpty
    )

    $bytes = [IO.File]::ReadAllBytes($NativePath)
    if ($bytes.Length -eq 0) {
        if (-not $AllowEmpty) {
            throw "Cannot convert an empty file: $NativePath"
        }
        $reviewParent = Split-Path -Parent $ReviewPath
        if (-not (Test-Path -LiteralPath $reviewParent)) {
            [void](New-Item -ItemType Directory -Path $reviewParent -Force)
        }
        Write-Utf8Text -Path $ReviewPath -Text ''
        return 'empty'
    }

    $encodingName = ''
    $text = $null

    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $encodingName = 'UTF-16 LE with BOM'
        $text = [Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $encodingName = 'UTF-16 BE with BOM'
        $text = [Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    elseif ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $encodingName = 'UTF-8 with BOM'
        $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
        $text = $strictUtf8.GetString($bytes, 3, $bytes.Length - 3)
    }
    else {
        $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
        try {
            $text = $strictUtf8.GetString($bytes)
            $encodingName = 'UTF-8 without BOM'
        }
        catch {
            $encoderFallback = New-Object System.Text.EncoderExceptionFallback
            $decoderFallback = New-Object System.Text.DecoderExceptionFallback
            $cp932 = [Text.Encoding]::GetEncoding(932, $encoderFallback, $decoderFallback)
            $text = $cp932.GetString($bytes)
            $encodingName = 'CP932 without BOM'
        }
    }

    if ($text.IndexOf([char]0) -ge 0) {
        throw "Decoded text contains an unexpected NUL character: $NativePath"
    }
    if ($text.IndexOf([char]0xFFFD) -ge 0) {
        throw "Decoded text contains a replacement character: $NativePath"
    }

    $reviewParent = Split-Path -Parent $ReviewPath
    if (-not (Test-Path -LiteralPath $reviewParent)) {
        [void](New-Item -ItemType Directory -Path $reviewParent -Force)
    }
    Write-Utf8Text -Path $ReviewPath -Text $text

    return $encodingName
}

function Protect-ConnectionString {
    param([AllowEmptyString()][string]$ConnectionString)

    if ([string]::IsNullOrWhiteSpace($ConnectionString)) {
        return ''
    }

    $safeKeys = @(
        'DRIVER', 'PROVIDER', 'TRUSTED_CONNECTION', 'INTEGRATED SECURITY',
        'ENCRYPT', 'TRUSTSERVERCERTIFICATE', 'MULTISUBNETFAILOVER',
        'MARS_CONNECTION', 'APPLICATIONINTENT', 'READONLY', 'HDR', 'IMEX',
        'FMT', 'CHARACTERSET'
    )

    try {
        $body = $ConnectionString.Trim()
        $prefix = ''
        if ($body -match '(?i)^ODBC\s*;') {
            $prefix = 'ODBC;'
            $body = $body.Substring($Matches[0].Length)
        }

        if ($prefix) {
            $parsed = New-Object System.Data.Odbc.OdbcConnectionStringBuilder
            $protected = New-Object System.Data.Odbc.OdbcConnectionStringBuilder
        }
        else {
            $parsed = New-Object System.Data.Common.DbConnectionStringBuilder
            $protected = New-Object System.Data.Common.DbConnectionStringBuilder
        }
        $parsed.set_ConnectionString($body)
        foreach ($keyObject in $parsed.Keys) {
            $key = [string]$keyObject
            if ([string]::IsNullOrWhiteSpace([string]$parsed[$key])) {
                continue
            }
            if ($safeKeys -contains $key.ToUpperInvariant()) {
                $protected[$key] = $parsed[$key]
            }
            else {
                $protected[$key] = '<redacted>'
            }
        }
        return $prefix + $protected.get_ConnectionString()
    }
    catch {
        return '<redacted-unparsed-connection-string>'
    }
}

function Get-AccessCollectionNames {
    param([Parameter(Mandatory = $true)][object]$Collection)

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($item in $Collection) {
        try {
            [void]$names.Add([string]$item.Name)
        }
        finally {
            Release-ComObject $item
        }
    }
    return @($names.ToArray() | Sort-Object)
}

function Get-QueryNames {
    param([Parameter(Mandatory = $true)][object]$Database)

    $names = New-Object System.Collections.Generic.List[string]
    $queryDefs = $null
    try {
        $queryDefs = $Database.QueryDefs
        foreach ($queryDef in $queryDefs) {
            try {
                $name = [string]$queryDef.Name
                if (-not $name.StartsWith('~') -and -not $name.StartsWith('MSys')) {
                    [void]$names.Add($name)
                }
            }
            finally {
                Release-ComObject $queryDef
            }
        }
    }
    finally {
        Release-ComObject $queryDefs
    }
    return @($names.ToArray() | Sort-Object)
}

function Get-LocalTableNames {
    param([Parameter(Mandatory = $true)][object]$Database)

    $names = New-Object System.Collections.Generic.List[string]
    $tableDefs = $null
    try {
        $tableDefs = $Database.TableDefs
        foreach ($tableDef in $tableDefs) {
            try {
                $name = [string]$tableDef.Name
                if (-not $name.StartsWith('MSys') -and
                    -not $name.StartsWith('~') -and
                    [string]::IsNullOrWhiteSpace([string]$tableDef.Connect)) {
                    [void]$names.Add($name)
                }
            }
            finally {
                Release-ComObject $tableDef
            }
        }
    }
    finally {
        Release-ComObject $tableDefs
    }
    return @($names.ToArray() | Sort-Object)
}

function Get-DatabasePropertyRecord {
    param(
        [Parameter(Mandatory = $true)][object]$Database,
        [Parameter(Mandatory = $true)][string]$Name,
        [object]$EffectiveWhenMissing = $null
    )

    $property = $null
    try {
        $property = $Database.Properties.Item($Name)
        return [pscustomobject]@{
            name = $Name
            defined = $true
            value = $property.Value
            effective_when_missing = $null
        }
    }
    catch {
        if (-not (Test-ExceptionErrorNumber -Exception $_.Exception -ExpectedNumbers @(3265, 3270))) {
            throw
        }
        return [pscustomobject]@{
            name = $Name
            defined = $false
            value = $null
            effective_when_missing = $EffectiveWhenMissing
        }
    }
    finally {
        Release-ComObject $property
    }
}

function Get-DatabaseSettingsFromCopy {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [SecureString]$Password
    )

    $engine = $null
    $databaseForSettings = $null
    $passwordPointer = [IntPtr]::Zero
    $plainTextPassword = $null
    try {
        $engine = New-Object -ComObject DAO.DBEngine.120
        if ($null -ne $Password) {
            $passwordPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
            $plainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordPointer)
            $databaseForSettings = $engine.OpenDatabase($Path, $false, $true, ";PWD=$plainTextPassword")
        }
        else {
            $databaseForSettings = $engine.OpenDatabase($Path, $false, $true)
        }

        return @(
            Get-DatabasePropertyRecord -Database $databaseForSettings -Name 'AllowBypassKey' -EffectiveWhenMissing $true
            Get-DatabasePropertyRecord -Database $databaseForSettings -Name 'StartupForm'
            Get-DatabasePropertyRecord -Database $databaseForSettings -Name 'StartupShowDBWindow' -EffectiveWhenMissing $true
            Get-DatabasePropertyRecord -Database $databaseForSettings -Name 'StartupShowStatusBar' -EffectiveWhenMissing $true
            Get-DatabasePropertyRecord -Database $databaseForSettings -Name 'AllowFullMenus' -EffectiveWhenMissing $true
            Get-DatabasePropertyRecord -Database $databaseForSettings -Name 'AllowBuiltinToolbars' -EffectiveWhenMissing $true
            Get-DatabasePropertyRecord -Database $databaseForSettings -Name 'AllowSpecialKeys' -EffectiveWhenMissing $true
        )
    }
    finally {
        if ($null -ne $databaseForSettings) {
            try { $databaseForSettings.Close() } catch {}
        }
        Release-ComObject $databaseForSettings
        Release-ComObject $engine
        if ($passwordPointer -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordPointer)
        }
        $plainTextPassword = $null
    }
}

function Get-TableMetadata {
    param(
        [Parameter(Mandatory = $true)][object]$Database,
        [switch]$ReadLinkedDetails
    )

    $records = New-Object System.Collections.Generic.List[object]
    $tableDefs = $null
    try {
        $tableDefs = $Database.TableDefs
        foreach ($tableDef in $tableDefs) {
            try {
                $name = [string]$tableDef.Name
                if ($name.StartsWith('MSys') -or $name.StartsWith('~')) {
                    continue
                }

                $connect = [string]$tableDef.Connect
                $isLinked = -not [string]::IsNullOrWhiteSpace($connect)
                $fields = New-Object System.Collections.Generic.List[object]
                $indexes = New-Object System.Collections.Generic.List[object]
                $detailsStatus = 'captured'

                if ($isLinked -and -not $ReadLinkedDetails) {
                    $detailsStatus = 'skipped-linked-table-details'
                }
                else {
                    $fieldCollection = $null
                    $indexCollection = $null
                    try {
                        $fieldCollection = $tableDef.Fields
                        foreach ($field in $fieldCollection) {
                            try {
                                [void]$fields.Add([pscustomobject]@{
                                    name = [string]$field.Name
                                    type = [int]$field.Type
                                    size = [int]$field.Size
                                    attributes = [long]$field.Attributes
                                    required = [bool]$field.Required
                                    allow_zero_length = [bool]$field.AllowZeroLength
                                    default_value = [string]$field.DefaultValue
                                    validation_rule = [string]$field.ValidationRule
                                })
                            }
                            finally {
                                Release-ComObject $field
                            }
                        }

                        $indexCollection = $tableDef.Indexes
                        foreach ($index in $indexCollection) {
                            $indexFields = New-Object System.Collections.Generic.List[string]
                            $indexFieldCollection = $null
                            try {
                                $indexFieldCollection = $index.Fields
                                foreach ($indexField in $indexFieldCollection) {
                                    try {
                                        [void]$indexFields.Add([string]$indexField.Name)
                                    }
                                    finally {
                                        Release-ComObject $indexField
                                    }
                                }
                                [void]$indexes.Add([pscustomobject]@{
                                    name = [string]$index.Name
                                    primary = [bool]$index.Primary
                                    unique = [bool]$index.Unique
                                    required = [bool]$index.Required
                                    ignore_nulls = [bool]$index.IgnoreNulls
                                    fields = @($indexFields.ToArray())
                                })
                            }
                            finally {
                                Release-ComObject $indexFieldCollection
                                Release-ComObject $index
                            }
                        }
                    }
                    catch {
                        $detailsStatus = 'error'
                        Add-ExportError -Stage 'table-metadata' -ObjectType 'Table' -ObjectName $name -Message $_.Exception.Message
                    }
                    finally {
                        Release-ComObject $indexCollection
                        Release-ComObject $fieldCollection
                    }
                }

                [void]$records.Add([pscustomobject]@{
                    name = $name
                    attributes = [long]$tableDef.Attributes
                    linked = $isLinked
                    source_table_name = [string]$tableDef.SourceTableName
                    connect = Protect-ConnectionString -ConnectionString $connect
                    details_status = $detailsStatus
                    fields = @($fields.ToArray())
                    indexes = @($indexes.ToArray())
                })
            }
            finally {
                Release-ComObject $tableDef
            }
        }
    }
    finally {
        Release-ComObject $tableDefs
    }

    return @($records.ToArray() | Sort-Object name)
}

function Get-RelationMetadata {
    param([Parameter(Mandatory = $true)][object]$Database)

    $records = New-Object System.Collections.Generic.List[object]
    $relations = $null
    try {
        $relations = $Database.Relations
        foreach ($relation in $relations) {
            $fields = New-Object System.Collections.Generic.List[object]
            $relationFields = $null
            try {
                $relationFields = $relation.Fields
                foreach ($field in $relationFields) {
                    try {
                        [void]$fields.Add([pscustomobject]@{
                            name = [string]$field.Name
                            foreign_name = [string]$field.ForeignName
                        })
                    }
                    finally {
                        Release-ComObject $field
                    }
                }
                [void]$records.Add([pscustomobject]@{
                    name = [string]$relation.Name
                    table = [string]$relation.Table
                    foreign_table = [string]$relation.ForeignTable
                    attributes = [long]$relation.Attributes
                    fields = @($fields.ToArray())
                })
            }
            finally {
                Release-ComObject $relationFields
                Release-ComObject $relation
            }
        }
    }
    finally {
        Release-ComObject $relations
    }
    return @($records.ToArray() | Sort-Object name)
}

function Get-ReferenceMetadata {
    param([Parameter(Mandatory = $true)][object]$AccessApplication)

    $records = New-Object System.Collections.Generic.List[object]
    $references = $null
    try {
        $references = $AccessApplication.References
        foreach ($reference in $references) {
            try {
                $fullPath = ''
                try { $fullPath = [string]$reference.FullPath } catch {}
                [void]$records.Add([pscustomobject]@{
                    name = [string]$reference.Name
                    guid = [string]$reference.Guid
                    major = [int]$reference.Major
                    minor = [int]$reference.Minor
                    broken = [bool]$reference.IsBroken
                    file_name = if ($fullPath) { [IO.Path]::GetFileName($fullPath) } else { '' }
                })
            }
            catch {
                Add-ExportError -Stage 'reference-metadata' -ObjectType 'Reference' -ObjectName '' -Message $_.Exception.Message
            }
            finally {
                Release-ComObject $reference
            }
        }
    }
    finally {
        Release-ComObject $references
    }
    return @($records.ToArray() | Sort-Object name)
}

function Write-MetadataFile {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][object]$Value
    )

    $path = Join-Path $script:NativeMetadataDir $Name
    Write-Utf8Json -Path $path -Value $Value
    Add-ManifestFile -Tier 'native' -SourceKind 'metadata' -Category 'metadata' -Path $path -Encoding 'UTF-8 without BOM'
}

function Export-ObjectCategory {
    param(
        [Parameter(Mandatory = $true)][object]$AccessApplication,
        [Parameter(Mandatory = $true)][string]$Category,
        [Parameter(Mandatory = $true)][string]$ObjectType,
        [Parameter(Mandatory = $true)][int]$AccessObjectType,
        [Parameter(Mandatory = $true)][string]$Extension,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Names
    )

    $nativeDir = Join-Path $script:NativeSaveAsTextDir $Category
    $reviewDir = Join-Path $script:ReviewSaveAsTextDir $Category
    [void](New-Item -ItemType Directory -Path $nativeDir -Force)
    [void](New-Item -ItemType Directory -Path $reviewDir -Force)

    $index = 0
    foreach ($name in $Names) {
        $index++
        $fileName = Get-SafeFileName -Index $index -Name $name -Extension $Extension
        $nativePath = Join-Path $nativeDir $fileName
        $reviewPath = Join-Path $reviewDir $fileName

        try {
            $AccessApplication.SaveAsText($AccessObjectType, $name, $nativePath)
            $allowEmpty = $Category -eq 'modules'
            $detectedEncoding = Convert-ToReviewUtf8 -NativePath $nativePath -ReviewPath $reviewPath -AllowEmpty:$allowEmpty
            $reviewEncoding = if ($detectedEncoding -eq 'empty') { 'empty' } else { 'UTF-8 without BOM' }
            Add-ManifestFile -Tier 'native' -SourceKind 'SaveAsText' -Category $Category -ObjectType $ObjectType -ObjectName $name -Path $nativePath -Encoding $detectedEncoding -AllowEmpty:$allowEmpty
            Add-ManifestFile -Tier 'review_utf8' -SourceKind 'derived-text' -Category $Category -ObjectType $ObjectType -ObjectName $name -Path $reviewPath -Encoding $reviewEncoding -AllowEmpty:$allowEmpty
            if ((Get-Item -LiteralPath $nativePath).Length -eq 0) {
                [void]$script:EmptyObjects.Add([pscustomobject]@{
                    category = $Category
                    object_type = $ObjectType
                    object_name = $name
                })
            }
            $script:Exported[$Category]++
        }
        catch {
            Add-ExportError -Stage 'save-as-text' -ObjectType $ObjectType -ObjectName $name -Message $_.Exception.Message
        }
    }
}

function Export-TableDataMacros {
    param(
        [Parameter(Mandatory = $true)][object]$AccessApplication,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$TableNames
    )

    $category = 'data_macros'
    $nativeDir = Join-Path $script:NativeSaveAsTextDir $category
    $reviewDir = Join-Path $script:ReviewSaveAsTextDir $category
    [void](New-Item -ItemType Directory -Path $nativeDir -Force)
    [void](New-Item -ItemType Directory -Path $reviewDir -Force)

    $index = 0
    foreach ($tableName in $TableNames) {
        $index++
        $script:DataMacroTablesProbed++
        $fileName = Get-SafeFileName -Index $index -Name $tableName -Extension '.xml'
        $nativePath = Join-Path $nativeDir $fileName
        $reviewPath = Join-Path $reviewDir $fileName

        try {
            $AccessApplication.SaveAsText(12, $tableName, $nativePath)
            $script:Discovered[$category]++
            $detectedEncoding = Convert-ToReviewUtf8 -NativePath $nativePath -ReviewPath $reviewPath
            Add-ManifestFile -Tier 'native' -SourceKind 'SaveAsText' -Category $category -ObjectType 'TableDataMacro' -ObjectName $tableName -Path $nativePath -Encoding $detectedEncoding
            Add-ManifestFile -Tier 'review_utf8' -SourceKind 'derived-text' -Category $category -ObjectType 'TableDataMacro' -ObjectName $tableName -Path $reviewPath -Encoding 'UTF-8 without BOM'
            $script:Exported[$category]++
        }
        catch {
            if (Test-ExceptionErrorNumber -Exception $_.Exception -ExpectedNumbers @(2950)) {
                $script:DataMacroAbsentCount++
                Remove-Item -LiteralPath $nativePath -Force -ErrorAction SilentlyContinue
                Remove-Item -LiteralPath $reviewPath -Force -ErrorAction SilentlyContinue
                continue
            }

            $script:DataMacroProbeErrorCount++
            Add-ExportError -Stage 'table-data-macro' -ObjectType 'TableDataMacro' -ObjectName $tableName -Message $_.Exception.Message
        }
    }
}

function Invoke-ReviewSecretScan {
    $pattern = [regex]'(?i)(?:^|[;"])\s*(PWD|PASSWORD|UID|USER ID|USERID|USER|USERNAME|ACCESS TOKEN|SERVER|SERVERNAME|INSTANCE|DATA SOURCE|ADDRESS|ADDR|NETWORK ADDRESS|DATABASE|INITIAL CATALOG|DSN|DBQ|HOST|PORT)\s*='
    $seen = New-Object System.Collections.Generic.HashSet[string]

    foreach ($record in @($script:Manifest | Where-Object { $_.tier -eq 'review_utf8' })) {
        $path = Join-Path $script:OutputPath ($record.relative_path.Replace('/', '\'))
        $rawText = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)
        $scanText = [regex]::Replace($rawText, '"\s*\r?\n\s*"', '')
        foreach ($match in $pattern.Matches($scanText)) {
            $key = ([string]$match.Groups[1].Value).ToUpperInvariant()
            $lineNumber = 1 + [regex]::Matches($scanText.Substring(0, $match.Index), '\r\n|\n').Count
            $dedupeKey = "$($record.relative_path)|$lineNumber|$key"
            if (-not $seen.Add($dedupeKey)) {
                continue
            }

            $hit = [pscustomobject]@{
                relative_path = $record.relative_path
                line = $lineNumber
                key = $key
                file_sha256 = $record.sha256
            }
            [void]$script:SensitiveTokenHits.Add($hit)
            if ($SensitiveOutputPolicy -eq 'Fail') {
                Add-ExportError -Stage 'secret-scan' -ObjectType $record.object_type -ObjectName $record.object_name -Message (
                    "Potential connection secret or endpoint key '$key' found in $($record.relative_path) near line $lineNumber. The value was not logged."
                )
            }
        }
    }

    if ($script:SensitiveTokenHits.Count -gt 0 -and $SensitiveOutputPolicy -eq 'RestrictedLocal') {
        Add-LogLine -Message "Sensitive source candidates detected: $($script:SensitiveTokenHits.Count). Output is restricted to the local isolated stage."
    }
}

function Start-ShiftWatchdog {
    param(
        [Parameter(Mandatory = $true)][int]$AccessPid,
        [Parameter(Mandatory = $true)][string]$ExpectedAccessExe,
        [Parameter(Mandatory = $true)][long]$ExpectedCreationTimeUtcTicks,
        [Parameter(Mandatory = $true)][int]$TimeoutSeconds,
        [Parameter(Mandatory = $true)][string]$MarkerPath
    )

    $watchdogPath = Join-Path $script:WorkingDir 'shift-watchdog.ps1'
    $watchdogSource = @'
param(
    [int]$AccessPid,
    [string]$ExpectedAccessExe,
    [long]$ExpectedCreationTimeUtcTicks,
    [int]$TimeoutSeconds,
    [string]$MarkerPath
)

$ErrorActionPreference = 'SilentlyContinue'
Start-Sleep -Seconds $TimeoutSeconds

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class ShiftReleaseNative
{
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    public static void ShiftUp() { keybd_event(0x10, 0, 0x0002, UIntPtr.Zero); }
}
"@

[ShiftReleaseNative]::ShiftUp()
$stopped = $false
$process = Get-CimInstance Win32_Process -Filter "ProcessId=$AccessPid"
if ($null -ne $process -and
    $process.Name -ieq 'MSACCESS.EXE' -and
    $process.ExecutablePath -ieq $ExpectedAccessExe -and
    $null -ne $process.CreationDate -and
    ([DateTime]$process.CreationDate).ToUniversalTime().Ticks -eq $ExpectedCreationTimeUtcTicks) {
    Stop-Process -Id $AccessPid -Force
    $stopped = $true
}

$payload = "watchdog_fired=true`naccess_pid=$AccessPid`nprocess_stopped=$stopped`n"
[IO.File]::WriteAllText($MarkerPath, $payload, (New-Object Text.UTF8Encoding($false)))
'@
    Write-Utf8Text -Path $watchdogPath -Text $watchdogSource

    $windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $windowsPowerShell)) {
        throw "Windows PowerShell was not found: $windowsPowerShell"
    }

    $arguments = @(
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $watchdogPath),
        '-AccessPid', $AccessPid,
        '-ExpectedAccessExe', ('"{0}"' -f $ExpectedAccessExe),
        '-ExpectedCreationTimeUtcTicks', $ExpectedCreationTimeUtcTicks,
        '-TimeoutSeconds', $TimeoutSeconds,
        '-MarkerPath', ('"{0}"' -f $MarkerPath)
    )

    return Start-Process -FilePath $windowsPowerShell -ArgumentList $arguments -WindowStyle Hidden -PassThru
}

function Stop-ShiftWatchdog {
    param([object]$WatchdogProcess)

    if ($null -eq $WatchdogProcess) {
        return
    }
    try {
        if (-not $WatchdogProcess.HasExited) {
            Stop-Process -Id $WatchdogProcess.Id -Force
            Wait-Process -Id $WatchdogProcess.Id -Timeout 10 -ErrorAction SilentlyContinue
        }
    }
    catch {
        Add-ExportError -Stage 'watchdog-cleanup' -Message $_.Exception.Message
    }
}

function Test-ExpectedAccessProcess {
    param(
        [Parameter(Mandatory = $true)][int]$ProcessId,
        [Parameter(Mandatory = $true)][string]$ExpectedPath,
        [Parameter(Mandatory = $true)][long]$ExpectedCreationTimeUtcTicks
    )

    $process = Get-CimInstance Win32_Process -Filter "ProcessId=$ProcessId" -ErrorAction SilentlyContinue
    return ($null -ne $process -and
        $process.Name -ieq 'MSACCESS.EXE' -and
        $process.ExecutablePath -ieq $ExpectedPath -and
        $null -ne $process.CreationDate -and
        ([DateTime]$process.CreationDate).ToUniversalTime().Ticks -eq $ExpectedCreationTimeUtcTicks)
}

function Assert-RestrictedLocalOutputPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    if ($fullPath.StartsWith('\\', [StringComparison]::Ordinal)) {
        throw 'RestrictedLocal output cannot use a UNC path.'
    }

    $root = [IO.Path]::GetPathRoot($fullPath)
    $drive = New-Object IO.DriveInfo($root)
    if ($drive.DriveType -ne [IO.DriveType]::Fixed) {
        throw "RestrictedLocal output requires a fixed local drive: $root"
    }

    $current = $fullPath
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "RestrictedLocal output cannot traverse a reparse point: $current"
            }
        }
        $parent = Split-Path -Parent $current
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) {
            break
        }
        $current = $parent
    }
}

if ($SensitiveOutputPolicy -eq 'RestrictedLocal' -and -not $AcknowledgeRestrictedOutput) {
    throw 'RestrictedLocal requires -AcknowledgeRestrictedOutput.'
}
if ($SensitiveOutputPolicy -eq 'Fail' -and $AcknowledgeRestrictedOutput) {
    throw '-AcknowledgeRestrictedOutput is valid only with -SensitiveOutputPolicy RestrictedLocal.'
}

$script:SourcePath = (Resolve-Path -LiteralPath $DatabasePath).Path
$script:OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
if ($SensitiveOutputPolicy -eq 'RestrictedLocal') {
    Assert-RestrictedLocalOutputPath -Path $script:OutputPath
}
$sourceExtension = [IO.Path]::GetExtension($script:SourcePath)
if ($sourceExtension -notin @('.accdb', '.mdb')) {
    throw "DatabasePath must be an ACCDB or MDB file: $sourceExtension"
}
if (Test-Path -LiteralPath $script:OutputPath) {
    throw "OutputDirectory already exists: $script:OutputPath"
}

[void](New-Item -ItemType Directory -Path $script:OutputPath)
$script:WorkingDir = Join-Path $script:OutputPath '_working'
$script:NativeDir = Join-Path $script:OutputPath 'native'
$script:NativeSaveAsTextDir = Join-Path $script:NativeDir 'saveastext'
$script:NativeMetadataDir = Join-Path $script:NativeDir 'metadata'
$script:ReviewDir = Join-Path $script:OutputPath 'review_utf8'
$script:ReviewSaveAsTextDir = Join-Path $script:ReviewDir 'saveastext'

foreach ($directory in @(
    $script:WorkingDir,
    $script:NativeDir,
    $script:NativeSaveAsTextDir,
    $script:NativeMetadataDir,
    $script:ReviewDir,
    $script:ReviewSaveAsTextDir
)) {
    [void](New-Item -ItemType Directory -Path $directory -Force)
}

$workingDatabase = Join-Path $script:WorkingDir ([IO.Path]::GetFileName($script:SourcePath))
$watchdogMarker = Join-Path $script:WorkingDir 'watchdog-fired.txt'
$startedAt = Get-Date
$sourceHashBefore = (Get-FileHash -LiteralPath $script:SourcePath -Algorithm SHA256).Hash
$sourceHashAfter = $null
$workingHash = $null
$access = $null
$currentProject = $null
$database = $null
$allForms = $null
$allReports = $null
$allModules = $null
$allMacros = $null
$watchdog = $null
$shiftIsDown = $false
$accessPid = 0
$accessExe = ''
$accessCreationTimeUtcTicks = 0L
$accessCreationTimeUtc = ''
$accessVersion = ''
$accessBuild = ''
$openFormsAfterOpen = $null
$plainPassword = $null
$passwordBstr = [IntPtr]::Zero
$databaseSettings = @()

Add-LogLine -Message 'Export started.'

try {
    $environmentFingerprint = Get-EnvironmentFingerprint
    Write-MetadataFile -Name 'environment.json' -Value $environmentFingerprint
    $missingRequiredEnvironment = @(
        'PATHEXT', 'COMSPEC', 'SystemRoot', 'ProgramFiles', 'CommonProgramFiles', 'TEMP'
    ) | Where-Object {
        -not $environmentFingerprint.variables[$_].present
    }
    if ($missingRequiredEnvironment.Count -gt 0) {
        throw "Required process environment variables are missing: $($missingRequiredEnvironment -join ', ')"
    }

    Copy-Item -LiteralPath $script:SourcePath -Destination $workingDatabase
    $workingHash = (Get-FileHash -LiteralPath $workingDatabase -Algorithm SHA256).Hash
    if ($workingHash -ne $sourceHashBefore) {
        throw 'The disposable copy hash does not match the source hash.'
    }
    Add-LogLine -Message 'Disposable copy created and hash verified.'

    $databaseSettings = @(Get-DatabaseSettingsFromCopy -Path $workingDatabase -Password $DatabasePassword)
    Write-MetadataFile -Name 'startup-preflight.json' -Value $databaseSettings
    $allowBypassKey = @($databaseSettings | Where-Object { $_.name -eq 'AllowBypassKey' } | Select-Object -First 1)
    if ($allowBypassKey.Count -ne 1) {
        throw 'AllowBypassKey preflight did not return exactly one record.'
    }
    if ($allowBypassKey[0].defined -and
        -not [Convert]::ToBoolean($allowBypassKey[0].value, [Globalization.CultureInfo]::InvariantCulture)) {
        throw 'AllowBypassKey is explicitly False; refusing to rely on virtual Shift.'
    }
    Add-LogLine -Message 'DAO startup preflight completed before Access.Application was opened.'

    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.AutomationSecurity = 3
    $accessVersion = [string]$access.Version
    $accessBuild = [string]$access.Build

    [uint32]$pidValue = 0
    $windowHandle = [IntPtr]([long]$access.hWndAccessApp())
    if ($windowHandle -eq [IntPtr]::Zero) {
        throw 'Access hWndAccessApp was zero; refusing an untracked COM open.'
    }
    [void][AccessExportNative]::GetWindowThreadProcessId($windowHandle, [ref]$pidValue)
    $accessPid = [int]$pidValue
    if ($accessPid -le 0) {
        throw 'Could not resolve the dedicated Access PID from hWndAccessApp.'
    }

    $accessProcess = Get-CimInstance Win32_Process -Filter "ProcessId=$accessPid"
    if ($null -eq $accessProcess -or $accessProcess.Name -ine 'MSACCESS.EXE') {
        throw 'The PID resolved from hWndAccessApp is not MSACCESS.EXE.'
    }
    $accessExe = [string]$accessProcess.ExecutablePath
    if ([string]::IsNullOrWhiteSpace($accessExe)) {
        throw 'Could not record the Access executable path.'
    }
    if ($null -eq $accessProcess.CreationDate) {
        throw 'Could not record the Access process creation time.'
    }
    $accessCreationDate = ([DateTime]$accessProcess.CreationDate).ToUniversalTime()
    $accessCreationTimeUtcTicks = $accessCreationDate.Ticks
    $accessCreationTimeUtc = $accessCreationDate.ToString('o')
    Add-LogLine -Message "Dedicated Access PID recorded: $accessPid"

    $watchdog = Start-ShiftWatchdog -AccessPid $accessPid -ExpectedAccessExe $accessExe -ExpectedCreationTimeUtcTicks $accessCreationTimeUtcTicks -TimeoutSeconds $OpenTimeoutSeconds -MarkerPath $watchdogMarker

    try {
        if (([AccessExportNative]::GetAsyncKeyState([AccessExportNative]::VK_SHIFT) -band 0x8000) -ne 0) {
            throw 'Shift is already pressed; refusing to inject or release keyboard state.'
        }
        [AccessExportNative]::ShiftDown()
        $shiftIsDown = $true

        if ($null -ne $DatabasePassword) {
            $passwordBstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($DatabasePassword)
            $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordBstr)
            $access.OpenCurrentDatabase($workingDatabase, $false, $plainPassword)
        }
        else {
            $access.OpenCurrentDatabase($workingDatabase, $false)
        }
    }
    finally {
        if ($shiftIsDown) {
            [AccessExportNative]::ShiftUp()
            $shiftIsDown = $false
        }
        Stop-ShiftWatchdog -WatchdogProcess $watchdog
        $watchdog = $null
        if ($passwordBstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr)
            $passwordBstr = [IntPtr]::Zero
            $plainPassword = $null
        }
    }

    if (Test-Path -LiteralPath $watchdogMarker) {
        throw 'The startup watchdog fired while OpenCurrentDatabase was blocked.'
    }

    $formsCollection = $null
    try {
        $formsCollection = $access.Forms
        $openFormsAfterOpen = [int]$formsCollection.Count
    }
    finally {
        Release-ComObject $formsCollection
    }
    Add-LogLine -Message "Database opened with virtual Shift; open form count: $openFormsAfterOpen"
    if ($openFormsAfterOpen -gt 0) {
        throw "Startup bypass failed: $openFormsAfterOpen form(s) were open after OpenCurrentDatabase."
    }

    $currentProject = $access.CurrentProject
    $database = $access.CurrentDb()
    $allForms = $currentProject.AllForms
    $allReports = $currentProject.AllReports
    $allModules = $currentProject.AllModules
    $allMacros = $currentProject.AllMacros

    $formNames = @(Get-AccessCollectionNames -Collection $allForms)
    $reportNames = @(Get-AccessCollectionNames -Collection $allReports)
    $moduleNames = @(Get-AccessCollectionNames -Collection $allModules)
    $macroNames = @(Get-AccessCollectionNames -Collection $allMacros)
    $queryNames = @(Get-QueryNames -Database $database)
    $localTableNames = @(Get-LocalTableNames -Database $database)

    $script:Discovered.forms = $formNames.Count
    $script:Discovered.reports = $reportNames.Count
    $script:Discovered.modules = $moduleNames.Count
    $script:Discovered.macros = $macroNames.Count
    $script:Discovered.queries = $queryNames.Count

    $startupFormSetting = @($databaseSettings | Where-Object { $_.name -eq 'StartupForm' } | Select-Object -First 1)
    $startupFormName = ''
    if ($startupFormSetting.Count -eq 1 -and $startupFormSetting[0].defined) {
        $startupFormName = [string]$startupFormSetting[0].value
    }

    $catalog = [pscustomobject]@{
        mode = $Mode
        autoexec_macro_present = @($macroNames | Where-Object { $_ -ieq 'AutoExec' }).Count -gt 0
        startup_form = $startupFormName
        forms = $formNames
        reports = $reportNames
        modules = $moduleNames
        macros = $macroNames
        queries = $queryNames
        local_tables = $localTableNames
    }

    if ($Mode -eq 'StartupProbe') {
        $reportNames = @()
        [void]$script:SkippedByMode.Add('reports')
    }

    Export-ObjectCategory -AccessApplication $access -Category 'forms' -ObjectType 'Form' -AccessObjectType 2 -Extension '.txt' -Names $formNames
    Export-ObjectCategory -AccessApplication $access -Category 'reports' -ObjectType 'Report' -AccessObjectType 3 -Extension '.txt' -Names $reportNames
    Export-ObjectCategory -AccessApplication $access -Category 'modules' -ObjectType 'Module' -AccessObjectType 5 -Extension '.txt' -Names $moduleNames
    Export-ObjectCategory -AccessApplication $access -Category 'macros' -ObjectType 'Macro' -AccessObjectType 4 -Extension '.txt' -Names $macroNames
    Export-ObjectCategory -AccessApplication $access -Category 'queries' -ObjectType 'Query' -AccessObjectType 1 -Extension '.txt' -Names $queryNames
    Export-TableDataMacros -AccessApplication $access -TableNames $localTableNames

    if ($Mode -eq 'Full') {
        $tables = @(Get-TableMetadata -Database $database -ReadLinkedDetails:$IncludeLinkedTableDetails)
        $relations = @(Get-RelationMetadata -Database $database)
        $references = @(Get-ReferenceMetadata -AccessApplication $access)
        Write-MetadataFile -Name 'tables.json' -Value $tables
        Write-MetadataFile -Name 'relations.json' -Value $relations
        Write-MetadataFile -Name 'references.json' -Value $references
    }
    Write-MetadataFile -Name 'catalog.json' -Value $catalog
    Write-MetadataFile -Name 'database-settings.json' -Value ([pscustomobject]@{
        access_version = $accessVersion
        access_build = $accessBuild
        properties = $databaseSettings
    })
    Invoke-ReviewSecretScan
}
catch {
    Add-ExportError -Stage 'fatal' -Message $_.Exception.Message
}
finally {
    if ($shiftIsDown) {
        [AccessExportNative]::ShiftUp()
        $shiftIsDown = $false
    }
    Stop-ShiftWatchdog -WatchdogProcess $watchdog

    if ($passwordBstr -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr)
        $passwordBstr = [IntPtr]::Zero
        $plainPassword = $null
    }

    Release-ComObject $allMacros
    Release-ComObject $allModules
    Release-ComObject $allReports
    Release-ComObject $allForms
    Release-ComObject $database
    Release-ComObject $currentProject

    if ($null -ne $access) {
        try { $access.CloseCurrentDatabase() } catch {}
        try { $access.Quit(2) } catch {}
    }
    Release-ComObject $access

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

try {
    $sourceHashAfter = (Get-FileHash -LiteralPath $script:SourcePath -Algorithm SHA256).Hash
    if ($sourceHashAfter -ne $sourceHashBefore) {
        Add-ExportError -Stage 'source-integrity' -Message 'The source database hash changed during export.'
    }

    foreach ($category in @('forms', 'reports', 'modules', 'macros', 'queries', 'data_macros')) {
        if ($script:SkippedByMode -contains $category) {
            continue
        }
        if ($script:Discovered[$category] -ne $script:Exported[$category]) {
            Add-ExportError -Stage 'count-validation' -ObjectType $category -Message (
                "Discovered $($script:Discovered[$category]) but exported $($script:Exported[$category])."
            )
        }
    }

    foreach ($record in $script:Manifest) {
        $recordPath = Join-Path $script:OutputPath ($record.relative_path.Replace('/', '\'))
        if (-not (Test-Path -LiteralPath $recordPath)) {
            Add-ExportError -Stage 'manifest-validation' -Message "Manifest file is missing: $($record.relative_path)"
            continue
        }
        if ((Get-FileHash -LiteralPath $recordPath -Algorithm SHA256).Hash -ne $record.sha256) {
            Add-ExportError -Stage 'manifest-validation' -Message "Manifest hash mismatch: $($record.relative_path)"
        }
    }

    if ($accessPid -gt 0) {
        try {
            Wait-Process -Id $accessPid -Timeout 15 -ErrorAction SilentlyContinue
        }
        catch {}

        if (Test-ExpectedAccessProcess -ProcessId $accessPid -ExpectedPath $accessExe -ExpectedCreationTimeUtcTicks $accessCreationTimeUtcTicks) {
            Add-ExportError -Stage 'process-cleanup' -Message "The dedicated Access PID remained after Quit: $accessPid"
            Stop-Process -Id $accessPid -Force
            Wait-Process -Id $accessPid -Timeout 15 -ErrorAction SilentlyContinue
        }
    }

    $lockFiles = @(
        Get-ChildItem -LiteralPath $script:WorkingDir -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in @('.laccdb', '.ldb') }
    )
    if ($lockFiles.Count -gt 0) {
        Add-ExportError -Stage 'lock-cleanup' -Message 'A lock file remains beside the disposable database copy.'
    }
}
catch {
    Add-ExportError -Stage 'post-validation' -Message $_.Exception.Message
}

$manifestPath = Join-Path $script:OutputPath 'manifest.json'
$manifestCsvPath = Join-Path $script:OutputPath 'manifest.csv'
$errorsPath = Join-Path $script:OutputPath 'export-errors.json'
$logPath = Join-Path $script:OutputPath 'export.log'
$summaryPath = Join-Path $script:OutputPath 'export-summary.json'
$sensitiveFindingsPath = Join-Path $script:OutputPath 'sensitive-findings.json'

Write-Utf8Json -Path $manifestPath -Value @($script:Manifest.ToArray())
$script:Manifest.ToArray() | Export-Csv -LiteralPath $manifestCsvPath -NoTypeInformation -Encoding UTF8
Write-Utf8Json -Path $errorsPath -Value @($script:Errors.ToArray())
Write-Utf8Text -Path $logPath -Text (($script:LogLines.ToArray() -join [Environment]::NewLine) + [Environment]::NewLine)
$sensitiveFindings = [ordered]@{
    schema_version = 1
    source_sha256 = $sourceHashBefore
    findings = @($script:SensitiveTokenHits.ToArray() | Sort-Object relative_path, line, key)
}
Write-Utf8Json -Path $sensitiveFindingsPath -Value $sensitiveFindings

$manifestHash = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
$sensitiveFindingsHash = (Get-FileHash -LiteralPath $sensitiveFindingsPath -Algorithm SHA256).Hash
$status = if ($script:Errors.Count -gt 0) {
    'FAIL'
}
elseif ($script:SensitiveTokenHits.Count -gt 0) {
    'PASS_RESTRICTED'
}
else {
    'PASS'
}
$summary = [ordered]@{
    status = $status
    mode = $Mode
    source_file = [IO.Path]::GetFileName($script:SourcePath)
    source_sha256_before = $sourceHashBefore
    source_sha256_after = $sourceHashAfter
    disposable_copy_sha256_before_open = $workingHash
    exporter_sha256 = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
    manifest_sha256 = $manifestHash
    access_version = $accessVersion
    access_build = $accessBuild
    startup_bypass_method = 'virtual-shift-with-external-watchdog'
    startup_absence_evidence = 'database-specific evidence is still required by the workflow'
    open_forms_after_open = $openFormsAfterOpen
    access_pid = $accessPid
    access_creation_time_utc = $accessCreationTimeUtc
    linked_table_details_requested = [bool]$IncludeLinkedTableDetails
    sensitive_output_policy = $SensitiveOutputPolicy
    restricted_output_acknowledged = [bool]$AcknowledgeRestrictedOutput
    disclosure_status = if ($script:SensitiveTokenHits.Count -gt 0) { 'RESTRICTED' } else { 'NO_CANDIDATE_DETECTED' }
    skipped_by_mode = @($script:SkippedByMode.ToArray())
    data_macro_tables_probed = $script:DataMacroTablesProbed
    data_macro_absent_count = $script:DataMacroAbsentCount
    data_macro_probe_error_count = $script:DataMacroProbeErrorCount
    sensitive_token_hit_count = $script:SensitiveTokenHits.Count
    sensitive_findings_sha256 = $sensitiveFindingsHash
    empty_object_count = $script:EmptyObjects.Count
    empty_objects = @($script:EmptyObjects.ToArray())
    discovered = $script:Discovered
    exported = $script:Exported
    manifest_records = $script:Manifest.Count
    error_count = $script:Errors.Count
    started_at = $startedAt.ToString('o')
    finished_at = (Get-Date).ToString('o')
}
Write-Utf8Json -Path $summaryPath -Value $summary

if ($status -in @('PASS', 'PASS_RESTRICTED') -and -not $KeepWorkingCopy) {
    $resolvedWorkingDir = [IO.Path]::GetFullPath($script:WorkingDir)
    $expectedWorkingDir = [IO.Path]::GetFullPath((Join-Path $script:OutputPath '_working'))
    $outputPrefix = $script:OutputPath.TrimEnd('\') + '\'
    if ($resolvedWorkingDir -ne $expectedWorkingDir -or
        -not $resolvedWorkingDir.StartsWith($outputPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Refusing to remove a working directory outside the verified output directory.'
    }
    Remove-Item -LiteralPath $resolvedWorkingDir -Recurse -Force
}

if ($status -eq 'FAIL') {
    throw "Access export failed. Review export-errors.json in: $script:OutputPath"
}

if ($status -eq 'PASS_RESTRICTED') {
    Write-Warning 'PASS_RESTRICTED: export is complete, but sensitive source candidates were detected. Keep every artifact in the local isolated stage and do not send it to AI or publish it.'
}
else {
    Write-Host "PASS: exported Access assets from a disposable copy."
}
Write-Host "Output: $script:OutputPath"
Write-Host "Manifest SHA-256: $manifestHash"
