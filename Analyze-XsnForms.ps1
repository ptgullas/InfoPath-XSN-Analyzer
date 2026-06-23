<#
.SYNOPSIS
    Inventory the logic, data, dependencies, views and complexity of InfoPath XSN form
    templates so a human can decide how to migrate them (Power Apps, Lists, rebuild, retire).

.DESCRIPTION
    READ-ONLY analysis tool. No parameters. Run it from the folder that contains the .xsn files
    (or drop it next to them). For each form you select it will:
      1. Extract the .xsn (a CAB) into a sibling folder named after the file (uses expand.exe).
      2. Parse manifest.xsf, the active view .xsl files, and the active .xsd schemas.
      3. Emit ONE multi-sheet .xlsx per form (via the ImportExcel module) plus a short .txt overview.
      4. Move orphaned / superseded files into a _Unused subfolder of the extracted folder.
    A roll-up workbook (_AllForms-Summary.xlsx) with one row per form + a complexity score is
    written at the end.

    The script never writes to SharePoint and never modifies the .xsn files. The only files it
    moves are the junk files inside each freshly-extracted folder.

    Designed to tolerate forms downloaded from SharePoint Server on-prem AND SharePoint Online,
    list-backed forms AND form-library (XML) forms. Unknown structures are reported, not fatal.

.NOTES
    Target: Windows PowerShell 5.1. Requires expand.exe (built into Windows) and the ImportExcel
    module (Install-Module ImportExcel -Scope CurrentUser).
#>

# ============================================================================================
#  Globals / setup
# ============================================================================================

$ErrorActionPreference = 'Stop'

# Resolve script root AFTER any binding (never in a param default). No params here, but keep the
# robust resolution so the script works when launched via the .bat (powershell.exe -File ...).
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Path }
if (-not $scriptRoot) { $scriptRoot = (Get-Location).Path }

# The people-picker / Contact Selector ActiveX control InfoPath stamps into list forms.
$ContactSelectorClsid = '61e40d31-993d-4777-8fa0-19ca59b6d0bb'

# SharePoint auto-managed columns - present on every list form, not author-defined. We label
# these 'System' rather than 'Unused' so the "candidates to drop" signal stays meaningful.
$SystemFields = @(
    'ID', 'Author', 'Editor', 'Modified', 'Created', 'Attachments', 'ContentType', 'ContentTypeId',
    'Order', 'owshiddenversion', 'Version', '_ModerationStatus', '_UIVersionString', 'FileLeafRef',
    'GUID', 'WorkflowVersion', 'InstanceID', 'ServerUrl', 'EncodedAbsUrl'
)

# InfoPath structural wrapper elements (data-source roots / groups) - not real fields.
$StructuralWrappers = @(
    'myFields', 'dataFields', 'queryFields', 'SharePointListItem_RW', 'Fields', 'Item', 'gpx'
)

# Element local-names we know how to describe. Anything outside these sets is surfaced as an
# "unrecognized" item (graceful degradation for forms built by other tools / newer InfoPath).
$KnownAdapters = @(
    'sharepointListAdapterRW', 'sharePointListAdapter', 'adoAdapter', 'webServiceAdapter',
    'davAdapter', 'xmlFileAdapter', 'emailAdapter', 'submitToHostAdapter', 'bdcAdapter',
    'grooveAdapter', 'hwsAdapter'
)
$KnownActions = @(
    'assignmentAction', 'switchViewAction', 'queryAction', 'submitAction', 'dialogBoxMessageAction',
    'dialogBoxExpressionAction', 'closeDocumentAction', 'openNewDocumentAction', 'changeAdapterProperty',
    'signSignatureLineAction', 'webPartConnectionAction', 'exitRuleSet', 'ruleSetAction'
)

# Friendly names for InfoPath control types (xd:xctname) - so a field's control (rich text vs plain,
# date picker, attachment, repeating table, people picker) is visible for migration planning.
$ControlTypeMap = @{
    'PlainText' = 'Text box'; 'RichText' = 'Rich text'; 'dropdown' = 'Drop-down'; 'combobox' = 'Combo box'
    'multiselectlistbox' = 'Multi-select list'; 'ListBox' = 'List box'; 'CheckBox' = 'Check box'
    'OptionButton' = 'Option button'; 'Button' = 'Button'; 'RepeatingTable' = 'Repeating table'
    'RepeatingSection' = 'Repeating section'; 'Section' = 'Section'; 'ExpressionBox' = 'Calculated / display'
    'FileAttachment' = 'File attachment'; 'PeoplePicker' = 'People picker'; 'BulletedList' = 'Bulleted list'
    'ListItem_Plain' = 'List item'; 'RichTextBox' = 'Rich text'
}

# ----------------------------------------------------------------------------------------------
#  Run log - every failure (extraction, manifest parse, any analyzer) is recorded here AND shown,
#  so a partially-parsed form is never mistaken for a fully-parsed one. Written to _Analysis-Log.txt.
# ----------------------------------------------------------------------------------------------
$script:RunLog = New-Object System.Collections.Generic.List[object]

function Log-Issue {
    param([string]$Form, [string]$Stage, [string]$Message, [string]$Severity = 'Warning')
    $script:RunLog.Add([pscustomobject]@{ Form = $Form; Stage = $Stage; Severity = $Severity; Message = $Message })
    Write-Warning ("    [{0}] {1}: {2}" -f $Severity, $Stage, $Message)
}

# ============================================================================================
#  Small helpers
# ============================================================================================

function Get-Attr {
    param($Node, [string]$Name)
    if ($null -eq $Node) { return '' }
    $v = $null
    try { $v = $Node.GetAttribute($Name) } catch { $v = '' }
    if ($null -eq $v) { return '' }
    return $v
}

function Sel {
    # SelectNodes that never throws; returns an array (possibly empty).
    param($Node, [string]$Xpath, $Nsm)
    if ($null -eq $Node) { return @() }
    $res = $null
    try { $res = $Node.SelectNodes($Xpath, $Nsm) } catch { return @() }
    if ($null -eq $res) { return @() }
    return @($res)
}

function Sel1 {
    param($Node, [string]$Xpath, $Nsm)
    if ($null -eq $Node) { return $null }
    $res = $null
    try { $res = $Node.SelectSingleNode($Xpath, $Nsm) } catch { return $null }
    return $res
}

function New-NsMgr {
    # Build a namespace manager seeded with the well-known InfoPath prefixes, then add every
    # xmlns:* declared on the document root (this captures the per-form 'my', 'd', 'q' URIs which
    # differ between list forms and library forms, and between on-prem and SPO).
    param([xml]$Doc)
    # NOTE: XmlNamespaceManager.AddNamespace() leaks a value into the PowerShell output
    # stream, and the manager is itself IEnumerable. Both can turn the function's return value
    # into an Object[]. Suppress every call with [void] and return with a unary comma so the
    # caller always gets the manager itself, never an array.
    $nsm = New-Object System.Xml.XmlNamespaceManager($Doc.NameTable)
    $fixed = @{
        'xsf'  = 'http://schemas.microsoft.com/office/infopath/2003/solutionDefinition'
        'xsf2' = 'http://schemas.microsoft.com/office/infopath/2006/solutionDefinition/extensions'
        'xsf3' = 'http://schemas.microsoft.com/office/infopath/2009/solutionDefinition/extensions'
        'xsd'  = 'http://www.w3.org/2001/XMLSchema'
        'xsl'  = 'http://www.w3.org/1999/XSL/Transform'
        'xd'   = 'http://schemas.microsoft.com/office/infopath/2003'
        'dfs'  = 'http://schemas.microsoft.com/office/infopath/2003/dataFormSolution'
    }
    foreach ($k in $fixed.Keys) {
        if (-not $nsm.HasNamespace($k)) { [void]$nsm.AddNamespace($k, $fixed[$k]) }
    }
    $root = $Doc.DocumentElement
    if ($null -ne $root) {
        foreach ($a in $root.Attributes) {
            if ($a.Prefix -eq 'xmlns') {
                $p = $a.LocalName
                if (-not $nsm.HasNamespace($p)) {
                    try { [void]$nsm.AddNamespace($p, $a.Value) } catch {}
                }
            }
        }
    }
    return ,$nsm
}

function Load-Xml {
    param([string]$Path)
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $false
    $doc.XmlResolver = $null
    $doc.Load($Path)
    return $doc
}

function Decode-Xml {
    param([string]$s)
    if ($null -eq $s) { return '' }
    $s = $s -replace '&lt;', '<' -replace '&gt;', '>' -replace '&quot;', '"' -replace '&apos;', "'" -replace '&amp;', '&'
    return $s
}

function Leaf-Name {
    # Last path segment of an XPath/binding, namespace prefix stripped: ../my:Foo/my:Bar -> Bar
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = $Path.Trim()
    $p = $p -replace '\[[^\]]*\]', ''          # drop predicates
    $segs = $p -split '/'
    $last = ''
    for ($i = $segs.Length - 1; $i -ge 0; $i--) {
        $seg = $segs[$i].Trim()
        if ($seg -ne '' -and $seg -ne '.' -and $seg -ne '..') { $last = $seg; break }
    }
    if ($last -match ':') { $last = ($last -split ':')[-1] }
    return $last
}

function To-PlainExpr {
    # Best-effort, human-readable rendering of an InfoPath XPath expression.
    param([string]$Expr)
    if ([string]::IsNullOrWhiteSpace($Expr)) { return '' }
    $s = Decode-Xml $Expr
    $s = $s -replace 'xdUser:get-UserName\(\)', 'current-user'
    $s = $s -replace 'xdEnvironment:[A-Za-z\-]+\(\)', 'environment'
    $s = $s -replace 'xdMath:Nz', 'numOrZero'
    $s = $s -replace 'xdDate:Today\(\)', 'today'
    $s = $s -replace 'xdDate:Now\(\)', 'now'
    # Strip namespace prefixes on node references (my:, d:, q:, dfs:) and path noise.
    $s = $s -replace '(\.\./)+', '' -replace '\./', ''
    $s = $s -replace '\b(my|d|q|dfs|pc|ma|dms):', ''
    # Operators to words (spaces guard against touching function names).
    $s = $s -replace '\s!=\s', ' is not ' -replace '\s=\s', ' is ' -replace '\s&gt;=\s', ' >= ' -replace '\s&lt;=\s', ' <= '
    $s = $s -replace '\sand\s', ' AND ' -replace '\sor\s', ' OR '
    $s = $s.Trim()
    if ($s.Length -gt 400) { $s = $s.Substring(0, 397) + '...' }
    return $s
}

function To-PowerFx {
    # Rough Power Fx visibility translation of an InfoPath XPath test. Clearly a starting point.
    param([string]$Expr)
    if ([string]::IsNullOrWhiteSpace($Expr)) { return '' }
    $s = Decode-Xml $Expr
    $s = $s -replace '(\.\./)+', '' -replace '\./', ''
    $s = $s -replace '\b(my|d|q|dfs|pc|ma|dms):', ''
    $s = $s -replace '!=', '<>'
    $s = $s -replace '\bnot\(', 'Not('
    $s = $s -replace '\band\b', 'And' -replace '\bor\b', 'Or'
    $s = $s -replace 'string-length\(', 'Len('
    $s = $s -replace 'xdUser:get-UserName\(\)', 'User().Email'
    $s = $s.Trim()
    if ($s.Length -gt 400) { $s = $s.Substring(0, 397) + '...' }
    return $s
}

function Clamp {
    param([string]$s, [int]$Max = 250)
    if ($null -eq $s) { return '' }
    if ($s.Length -gt $Max) { return $s.Substring(0, $Max - 3) + '...' }
    return $s
}

# ============================================================================================
#  Extraction
# ============================================================================================

function Get-FileHead {
    # First N bytes of a file (for archive-format sniffing).
    param([string]$Path, [int]$Count = 8)
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        try {
            $b = New-Object byte[] $Count
            $n = $fs.Read($b, 0, $Count)
            if ($n -lt $Count -and $n -gt 0) { $b = $b[0..($n - 1)] }
            elseif ($n -le 0) { return @() }
            return $b
        } finally { $fs.Close() }
    } catch { return @() }
}

function Extract-CabTo {
    # Extract a CAB (MSCF) to $DestDir using a temp .cab copy + expand.exe, with the Windows shell
    # cabinet handler as a fallback. Returns expand's text output.
    param([string]$SrcPath, [string]$DestDir, [string]$FormName)
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ($FormName + '_' + ([System.IO.Path]::GetRandomFileName().Replace('.', '')) + '.cab')
    $manifest = Join-Path $DestDir 'manifest.xsf'
    $out = ''
    try { Copy-Item -LiteralPath $SrcPath -Destination $tmp -Force } catch {}
    if (Test-Path -LiteralPath $tmp) {
        try { $out = & expand.exe $tmp '-F:*' $DestDir 2>&1 | Out-String } catch { $out = $_.Exception.Message }
    }
    if (-not (Test-Path -LiteralPath $manifest) -and (Test-Path -LiteralPath $tmp)) {
        try {
            $shell = New-Object -ComObject Shell.Application
            $s = $shell.NameSpace([string]$tmp); $d = $shell.NameSpace([string]$DestDir)
            if ($s -and $d) {
                $d.CopyHere($s.Items(), 0x14)
                for ($i = 0; $i -lt 75; $i++) { if (Test-Path -LiteralPath $manifest) { break }; Start-Sleep -Milliseconds 200 }
            }
        } catch {}
    }
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath (Join-Path $DestDir ([System.IO.Path]::GetFileName($tmp))) -Force -ErrorAction SilentlyContinue
    return $out
}

function Extract-ZipTo {
    # Extract a ZIP (PK..) to $DestDir using a temp .zip copy + Expand-Archive.
    param([string]$SrcPath, [string]$DestDir, [string]$FormName)
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ($FormName + '_' + ([System.IO.Path]::GetRandomFileName().Replace('.', '')) + '.zip')
    $out = ''
    try {
        Copy-Item -LiteralPath $SrcPath -Destination $tmp -Force
        Expand-Archive -LiteralPath $tmp -DestinationPath $DestDir -Force -ErrorAction Stop
    } catch { $out = $_.Exception.Message }
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    return $out
}

function Flatten-Manifest {
    # If manifest.xsf landed in a subfolder (some archives nest), copy that folder's files up to root.
    param([string]$DestDir)
    $manifest = Join-Path $DestDir 'manifest.xsf'
    if (Test-Path -LiteralPath $manifest) { return }
    $found = Get-ChildItem -LiteralPath $DestDir -Recurse -Filter 'manifest.xsf' -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) {
        try {
            Get-ChildItem -LiteralPath $found.DirectoryName -File -ErrorAction SilentlyContinue | ForEach-Object {
                Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $DestDir $_.Name) -Force
            }
        } catch {}
    }
}

function Expand-Xsn {
    param([string]$XsnPath, [string]$DestDir)
    if (-not (Test-Path -LiteralPath $DestDir)) {
        New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    }
    $manifest = Join-Path $DestDir 'manifest.xsf'
    if (Test-Path -LiteralPath $manifest) {
        return $true   # already extracted; idempotent re-run
    }
    $formName = [System.IO.Path]::GetFileNameWithoutExtension($XsnPath)

    # Do NOT gate on the file header. A .xsn re-saved by InfoPath "Save As" can be a valid cabinet
    # whose MSCF signature is not at byte 0 (renaming to .cab and extracting still works), so a strict
    # header check is a false negative. Just attempt extraction: CAB first (copy to .cab + expand +
    # shell handler, i.e. exactly what a manual .cab rename does), then ZIP as a fallback.
    $out1 = Extract-CabTo -SrcPath $XsnPath -DestDir $DestDir -FormName $formName
    Flatten-Manifest $DestDir
    if (Test-Path -LiteralPath $manifest) { return $true }

    $out2 = Extract-ZipTo -SrcPath $XsnPath -DestDir $DestDir -FormName $formName
    Flatten-Manifest $DestDir
    if (Test-Path -LiteralPath $manifest) { return $true }

    # Genuinely could not extract. Report the header bytes so the real format is identifiable.
    $head = Get-FileHead $XsnPath 8
    $hex = (@($head) | ForEach-Object { $_.ToString('X2') }) -join ' '
    $asc = -join (@($head) | ForEach-Object { if ($_ -ge 32 -and $_ -lt 127) { [char]$_ } else { '.' } })
    if (-not $hex) { $hex = '(empty / unreadable - likely a OneDrive cloud-only stub; hydrate it and re-run)' }
    Log-Issue $formName 'extract' "Could not extract as CAB or ZIP. First bytes: [$hex] '$asc'. cab=$(Clamp $out1 120) zip=$(Clamp $out2 120)" 'Error'
    return $false
}

# ============================================================================================
#  Manifest analysis
# ============================================================================================

function Get-FormType {
    param($Root, $Nsm)
    $name = Get-Attr $Root 'name'
    $mode = ''
    $isListEdit = ''
    $sm = Sel1 $Root '//xsf3:solutionMode' $Nsm
    if ($sm) { $mode = Get-Attr $sm 'mode'; $isListEdit = Get-Attr $sm 'isListEditForm' }
    $hasListAdapter = (Sel $Root '//xsf:sharepointListAdapterRW' $Nsm).Count -gt 0
    $hasDav = (Sel $Root '//xsf:davAdapter' $Nsm).Count -gt 0

    # Host list / location
    $baseUrl = ''
    $bu = Sel1 $Root '//xsf3:baseUrl' $Nsm
    if ($bu) { $baseUrl = Get-Attr $bu 'relativeUrlBase' }
    if (-not $baseUrl) { $baseUrl = Get-Attr $Root 'publishUrl' }

    $isList = ($name -match ':list:') -or ($mode -eq 'list') -or $hasListAdapter
    $isLibrary = $hasDav -or ($baseUrl -match '(?i)/Forms/') -or ($mode -eq 'wss')
    if ($isList) {
        $kind = 'SharePoint List form'
        $desc = 'Edits items in a SharePoint list (column-backed). Each submission is a list item.'
    } elseif ($isLibrary) {
        $kind = 'Form Library form (XML)'
        $desc = 'Saves an XML document into a SharePoint form library. Each submission is an .xml file.'
    } else {
        $kind = 'InfoPath form (XML)'
        $desc = 'Standalone / XML-document form. No SharePoint list adapter detected.'
    }
    return [pscustomobject]@{
        Kind = $kind; Description = $desc; SolutionName = $name; Mode = $mode
        IsListEditForm = $isListEdit; BaseUrl = $baseUrl
    }
}

function Get-Adapters {
    # Returns objects describing every data adapter (primary + secondary). Passwords redacted.
    param($Root, $Nsm)
    $rows = @()

    # Primary query adapter(s): direct children of xDocumentClass (not inside dataObjects).
    $primaryQ = Sel $Root '/xsf:xDocumentClass/xsf:query/*' $Nsm
    foreach ($a in $primaryQ) {
        $rows += (New-AdapterRow -Adapter $a -DataObjectName 'Main Data Connection' -IsPrimary $true -InitOnLoad $false)
    }
    # Inline submit adapter(s) - only real adapters (not useQueryAdapter / useScriptHandler etc.)
    $primaryS = Sel $Root '/xsf:xDocumentClass/xsf:submit/*' $Nsm
    foreach ($a in $primaryS) {
        if ($script:KnownAdapters -contains $a.LocalName) {
            $rows += (New-AdapterRow -Adapter $a -DataObjectName 'Submit connection' -IsPrimary $true -InitOnLoad $false)
        }
    }
    # Top-level named dataAdapters (some forms declare adapters here, referenced by rules)
    foreach ($a in (Sel $Root '/xsf:xDocumentClass/xsf:dataAdapters/*' $Nsm)) {
        if ($script:KnownAdapters -contains $a.LocalName) {
            $rows += (New-AdapterRow -Adapter $a -DataObjectName (Get-Attr $a 'name') -IsPrimary $false -InitOnLoad $false)
        }
    }
    # Secondary data objects
    $dos = Sel $Root '//xsf:dataObjects/xsf:dataObject' $Nsm
    foreach ($d in $dos) {
        $doName = Get-Attr $d 'name'
        $initOnLoad = (Get-Attr $d 'initOnLoad') -eq 'yes'
        $adapters = Sel $d './xsf:query/*' $Nsm
        foreach ($a in $adapters) {
            if ($script:KnownAdapters -contains $a.LocalName) {
                $rows += (New-AdapterRow -Adapter $a -DataObjectName $doName -IsPrimary $false -InitOnLoad $initOnLoad)
            }
        }
    }
    return $rows
}

function New-AdapterRow {
    param($Adapter, [string]$DataObjectName, [bool]$IsPrimary, [bool]$InitOnLoad)
    $ln = $Adapter.LocalName
    $kind = 'Other'
    $isSpList = $false
    $listId = ''; $listName = ''; $site = ''; $detail = ''
    $queryAllowed = Get-Attr $Adapter 'queryAllowed'
    $submitAllowed = Get-Attr $Adapter 'submitAllowed'

    if ($ln -eq 'sharepointListAdapterRW' -or $ln -eq 'sharePointListAdapter') {
        $kind = $(if ($ln -eq 'sharePointListAdapter') { 'SharePoint list (read-only)' } else { 'SharePoint list' })
        $isSpList = $true
        $listId = Get-Attr $Adapter 'sharePointListID'
        if (-not $listId) { $listId = Get-Attr $Adapter 'sharepointGuid' }
        $listName = Get-Attr $Adapter 'name'
        $site = Get-Attr $Adapter 'siteURL'
        if (-not $site) { $site = Get-Attr $Adapter 'siteUrl' }
        $rel = Get-Attr $Adapter 'relativeListUrl'
        $detail = "List '$listName' $listId site=$site rel=$rel"
    } elseif ($ln -eq 'adoAdapter') {
        $kind = 'SQL / database'
        $cs = Get-Attr $Adapter 'connectionString'
        $cmd = Get-Attr $Adapter 'commandText'
        # Redact any embedded credential before it ever reaches a report.
        $csRedacted = $cs -replace '(?i)(password|pwd)\s*=\s*[^;]*', '$1=***REDACTED***'
        $detail = "conn=[$($csRedacted)] cmd=[$(Clamp $cmd 160)]"
    } elseif ($ln -eq 'webServiceAdapter') {
        $kind = 'Web service (SOAP)'
        $wsdl = Get-Attr $Adapter 'wsdlUrl'
        $op = ''
        $opNode = $Adapter.SelectSingleNode('*[local-name()="operation"]')
        if ($opNode) { $op = $opNode.GetAttribute('name') }
        $detail = "wsdl=$wsdl op=$op"
    } elseif ($ln -eq 'davAdapter') {
        $kind = 'SharePoint library (submit)'
        $fn = ''
        $fnNode = $Adapter.SelectSingleNode('*[local-name()="fileName"]')
        if ($fnNode) { $fn = $fnNode.GetAttribute('value') }
        $detail = "library submit; fileName=$fn"
    } elseif ($ln -eq 'xmlFileAdapter') {
        $isRest = $false
        # A REST connection is an xmlFileAdapter flagged in the xsf2 extension.
        $kind = 'XML file'
        $detail = "file=$(Get-Attr $Adapter 'fileUrl')"
    } elseif ($ln -eq 'emailAdapter') {
        $kind = 'Email (submit)'
        $detail = "email submit"
    } elseif ($ln -eq 'submitToHostAdapter') {
        $kind = 'Submit to host page'
        $detail = "host submit (browser/SPFx hosting)"
    } elseif ($ln -eq 'bdcAdapter') {
        $kind = 'Business Connectivity Services (BCS)'
        $detail = "entity=$(Get-Attr $Adapter 'entityName') lob=$(Get-Attr $Adapter 'lobSystemInstance')"
    } elseif ($ln -eq 'grooveAdapter') {
        $kind = 'SharePoint Workspace / Groove (offline)'
        $detail = 'offline list adapter'
    } elseif ($ln -eq 'hwsAdapter') {
        $kind = 'Human Workflow Services (legacy)'
        $detail = 'HWS adapter'
    } else {
        $kind = "Other ($ln)"
        $detail = $ln
    }
    return [pscustomobject]@{
        DataObject = $DataObjectName; Kind = $kind; IsPrimary = $IsPrimary; IsSharePointList = $isSpList
        ListId = $listId; ListName = $listName; Site = $site; OnLoad = $InitOnLoad
        QueryAllowed = ($queryAllowed -eq 'yes'); SubmitAllowed = ($submitAllowed -eq 'yes')
        Detail = $detail
    }
}

function Get-PrimaryFields {
    # The form's own list columns, from the main (submit-capable) SharePoint list adapter.
    param($Root, $Nsm)
    $rows = @()
    $main = $null
    foreach ($a in (Sel $Root '//xsf:sharepointListAdapterRW' $Nsm)) {
        if ((Get-Attr $a 'submitAllowed') -eq 'yes') { $main = $a; break }
    }
    if ($null -eq $main) {
        # fall back to the first list adapter
        $first = Sel1 $Root '//xsf:sharepointListAdapterRW' $Nsm
        if ($first) { $main = $first }
    }
    $primaryListId = ''
    if ($main) {
        $primaryListId = Get-Attr $main 'sharePointListID'
        foreach ($f in (Sel $main './xsf:field' $Nsm)) {
            $t = Get-Attr $f 'type'
            $complex = @('Lookup', 'Choice', 'MultiChoice', 'User', 'MultiUser', 'Attachments') -contains $t
            $rows += [pscustomobject]@{
                Name = Get-Attr $f 'internalName'; Type = $t
                Required = ((Get-Attr $f 'required') -eq 'yes'); IsComplex = $complex
                LookupTarget = Get-Attr $f 'auxDomName'; Source = 'List column'
            }
        }
    }
    return [pscustomobject]@{ Fields = $rows; PrimaryListId = $primaryListId; HasListAdapter = ($null -ne $main) }
}

function Get-SchemaFields {
    # Walk the root schema's element tree (for library/XML forms, or to supplement list columns).
    param([string]$FolderDir, $Root, $Nsm)
    $rows = @()
    # Determine the root schema location: documentSchema rootSchema="yes", else first dataObject @schema.
    $schemaFile = ''
    $ds = Sel1 $Root '//xsf:documentSchemas/xsf:documentSchema[@rootSchema="yes"]' $Nsm
    if ($ds) {
        $loc = Get-Attr $ds 'location'        # "namespace file.xsd"
        if ($loc -match '\s') { $schemaFile = ($loc -split '\s+')[-1] } else { $schemaFile = $loc }
    }
    if (-not $schemaFile) { return $rows }
    $path = Join-Path $FolderDir $schemaFile
    if (-not (Test-Path -LiteralPath $path)) { return $rows }
    $xsd = $null
    try { $xsd = Load-Xml $path } catch { return $rows }
    $xnsm = New-NsMgr $xsd
    # Top-level elements and their immediate field descendants.
    foreach ($el in (Sel $xsd.DocumentElement '//xsd:element[@name]' $xnsm)) {
        $nm = Get-Attr $el 'name'
        if (-not $nm) { continue }
        $ty = Get-Attr $el 'type'
        $maxOccurs = Get-Attr $el 'maxOccurs'
        $repeating = ($maxOccurs -eq 'unbounded') -or ($maxOccurs -match '^[2-9]')
        $rows += [pscustomobject]@{
            Name = $nm; Type = $ty; Repeating = $repeating
            Nillable = ((Get-Attr $el 'nillable') -eq 'true'); Source = 'Schema (' + $schemaFile + ')'
        }
    }
    return $rows
}

function Get-PromotedFields {
    # Form fields promoted to SharePoint columns (property promotion), keyed by field leaf-name -> the
    # SharePoint column name. Present mainly on form-library (XML) forms; everything NOT here is stored
    # only inside the form's XML and is not a queryable/visible SharePoint column.
    param($Root, $Nsm)
    $map = @{}
    foreach ($f in (Sel $Root "//*[local-name()='field'][@columnName]" $Nsm)) {
        $leaf = Leaf-Name (Get-Attr $f 'node')
        if (-not $leaf) { $leaf = Get-Attr $f 'name' }
        $col = Get-Attr $f 'name'
        if ($leaf -and -not $map.ContainsKey($leaf)) { $map[$leaf] = $col }
    }
    return $map
}

function Get-RootSchemaFile {
    param($Root, $Nsm)
    $ds = Sel1 $Root '//xsf:documentSchemas/xsf:documentSchema[@rootSchema="yes"]' $Nsm
    if (-not $ds) { return '' }
    $loc = Get-Attr $ds 'location'
    if ($loc -match '\s') { return ($loc -split '\s+')[-1] } else { return $loc }
}

function Walk-StructEl {
    # Recursively walk an xsd:element (resolving ref= to its top-level definition) to build the
    # form's section/field tree. InfoPath schemas are ref-based: myFields -> refs to section
    # elements -> refs to field elements. Groups (with a complexType) are sections; leaves are fields.
    param($El, [int]$Depth, [string]$ParentPath, [bool]$RepeatingFromRef, $Defs, $Rows, $Nsm)
    if ($Depth -gt 15 -or $null -eq $El) { return }
    $maxOccurs = Get-Attr $El 'maxOccurs'
    $repeating = $RepeatingFromRef -or ($maxOccurs -eq 'unbounded') -or ($maxOccurs -match '^[2-9]')
    $ref = Get-Attr $El 'ref'
    if ($ref) {
        $local = Leaf-Name $ref
        $def = $Defs[$local]
        if ($def) {
            Walk-StructEl -El $def -Depth $Depth -ParentPath $ParentPath -RepeatingFromRef $repeating -Defs $Defs -Rows $Rows -Nsm $Nsm
        } else {
            $p = $(if ($ParentPath) { "$ParentPath/$local" } else { $local })
            $kind = $(if ($repeating) { 'Repeating field' } else { 'Field' })
            $Rows.Add([pscustomobject]@{ Depth = $Depth; Name = $local; Type = $ref; Kind = $kind; Path = $p }) | Out-Null
        }
        return
    }
    $name = Get-Attr $El 'name'
    if (-not $name) { return }
    $type = Get-Attr $El 'type'
    $ct = Sel1 $El './xsd:complexType' $Nsm
    $isGroup = ($null -ne $ct)
    $p = $(if ($ParentPath) { "$ParentPath/$name" } else { $name })
    $kind = 'Field'
    if ($isGroup -and $repeating) { $kind = 'Repeating section' }
    elseif ($isGroup) { $kind = 'Section' }
    elseif ($repeating) { $kind = 'Repeating field' }
    $Rows.Add([pscustomobject]@{ Depth = $Depth; Name = $name; Type = $type; Kind = $kind; Path = $p }) | Out-Null
    if ($isGroup) {
        foreach ($child in (Sel $ct './xsd:sequence/xsd:element|./xsd:all/xsd:element|./xsd:choice/xsd:element' $Nsm)) {
            Walk-StructEl -El $child -Depth ($Depth + 1) -ParentPath $p -RepeatingFromRef $false -Defs $Defs -Rows $Rows -Nsm $Nsm
        }
    }
}

function Get-FormStructure {
    # The section/field hierarchy of the form, from its root schema. Shows what sections the form
    # has and which fields live in each - the structural layout an LLM needs to rebuild it.
    param([string]$FolderDir, $Root, $Nsm)
    $rows = New-Object System.Collections.Generic.List[object]
    $schemaFile = Get-RootSchemaFile $Root $Nsm
    if (-not $schemaFile) { return $rows }
    $path = Join-Path $FolderDir $schemaFile
    if (-not (Test-Path -LiteralPath $path)) { return $rows }
    $xsd = $null
    try { $xsd = Load-Xml $path } catch { return $rows }
    $xnsm = New-NsMgr $xsd
    $topEls = @(Sel $xsd.DocumentElement './xsd:element' $xnsm)
    if ($topEls.Count -eq 0) { return $rows }
    $defs = @{}
    foreach ($e in $topEls) { $n = Get-Attr $e 'name'; if ($n) { $defs[$n] = $e } }
    $refNames = @{}
    foreach ($r in (Sel $xsd.DocumentElement '//xsd:element[@ref]' $xnsm)) { $rn = Leaf-Name (Get-Attr $r 'ref'); if ($rn) { $refNames[$rn] = $true } }
    $rootEl = $null
    foreach ($e in $topEls) { if (-not $refNames[(Get-Attr $e 'name')]) { $rootEl = $e; break } }
    if (-not $rootEl) { $rootEl = $topEls[0] }
    Walk-StructEl -El $rootEl -Depth 0 -ParentPath '' -RepeatingFromRef $false -Defs $defs -Rows $rows -Nsm $xnsm
    return $rows
}

function Friendly-Control {
    param([string]$Ct)
    if (-not $Ct) { return '' }
    if ($Ct -like 'DTPicker*') { return 'Date picker' }
    if ($script:ControlTypeMap.ContainsKey($Ct)) { return $script:ControlTypeMap[$Ct] }
    if ($Ct -match '^\{?[0-9a-fA-F-]{30,}\}?$') { return 'ActiveX control' }
    return $Ct
}

function Get-Defaults {
    # Static default values from the form template (initialXmlDocument / template.xml): field -> value.
    param([string]$FolderDir, $Root, $Nsm)
    $map = @{}
    $href = ''
    $init = Sel1 $Root '//xsf:fileNew/xsf:initialXmlDocument' $Nsm
    if ($init) { $href = Get-Attr $init 'href' }
    if (-not $href) { $href = 'template.xml' }
    $path = Join-Path $FolderDir $href
    if (-not (Test-Path -LiteralPath $path)) { return $map }
    $txt = ''
    try { $txt = [System.IO.File]::ReadAllText($path) } catch { return $map }
    # leaf elements only: <my:Field>text</my:Field> (groups contain child tags, so [^<]+ excludes them)
    foreach ($m in [regex]::Matches($txt, '<(?:my|d|q):(?<n>[A-Za-z0-9_]+)\b[^>]*>(?<v>[^<]+)</(?:my|d|q):\k<n>>')) {
        $n = $m.Groups['n'].Value
        $v = (Decode-Xml $m.Groups['v'].Value).Trim()
        if ($v -ne '' -and -not $map.ContainsKey($n)) { $map[$n] = $v }
    }
    return $map
}

function Build-RuleSetTriggers {
    # Map ruleSet name -> human trigger, by scanning onLoad and field-change handlers.
    param($Root, $Nsm)
    $map = @{}
    foreach ($r in (Sel $Root '//xsf:onLoad/xsf:ruleSetAction' $Nsm)) {
        $rs = Get-Attr $r 'ruleSet'
        if ($rs) { $map[$rs] = 'On form load' }
    }
    # Rules attached to the Submit event (newer forms store submit logic as a rule set).
    foreach ($r in (Sel $Root '//xsf:submit/xsf:ruleSetAction' $Nsm)) {
        $rs = Get-Attr $r 'ruleSet'
        if ($rs) { $map[$rs] = 'On submit' }
    }
    foreach ($h in (Sel $Root '//xsf:domEventHandlers/xsf:domEventHandler' $Nsm)) {
        $match = Get-Attr $h 'match'
        $leaf = Leaf-Name $match
        $trig = "On change of [$leaf]"
        if ($match -eq '/' -or $match -eq '.') { $trig = 'On form load' }
        foreach ($r in (Sel $h './xsf:ruleSetAction' $Nsm)) {
            $rs = Get-Attr $r 'ruleSet'
            if ($rs) { $map[$rs] = $trig }
        }
    }
    return ,$map
}

function Describe-Action {
    param($ActionNode)
    $ln = $ActionNode.LocalName
    if ($ln -eq 'assignmentAction') {
        $rawTgt = Get-Attr $ActionNode 'targetField'
        $tgt = Leaf-Name $rawTgt
        if (-not $tgt) { $tgt = 'this field' }
        $expr = To-PlainExpr (Get-Attr $ActionNode 'expression')
        if ($expr -eq '' -or $expr -eq '""') { return "Clear [$tgt]" }
        return "Set [$tgt] = $expr"
    } elseif ($ln -eq 'switchViewAction') {
        return "Go to view '" + (Get-Attr $ActionNode 'view') + "'"
    } elseif ($ln -eq 'submitAction') {
        return "Submit the form"
    } elseif ($ln -eq 'closeDocumentAction') {
        return "Close the form"
    } elseif ($ln -eq 'queryAction') {
        $adp = Get-Attr $ActionNode 'adapter'
        if ($adp) { return "Refresh connection '$adp'" }
        return "Run a query"
    } elseif ($ln -eq 'dialogBoxMessageAction') {
        return "Show message: " + (Clamp (Decode-Xml $ActionNode.InnerText) 120)
    } elseif ($ln -eq 'dialogBoxExpressionAction') {
        return "Show message: " + (To-PlainExpr (Get-Attr $ActionNode 'expression'))
    } elseif ($ln -eq 'openNewDocumentAction') {
        return "Open a new form from '" + (Get-Attr $ActionNode 'solutionURI') + "'"
    } elseif ($ln -eq 'changeAdapterProperty') {
        return "Change connection '" + (Get-Attr $ActionNode 'adapter') + "' property " + (Get-Attr $ActionNode 'adapterProperty')
    } elseif ($ln -eq 'signSignatureLineAction') {
        return "Sign a signature line (digital signature)"
    } elseif ($ln -eq 'webPartConnectionAction') {
        return "Send data over a Web Part connection"
    } elseif ($ln -eq 'ruleSetAction') {
        return "Run rule set '" + (Get-Attr $ActionNode 'ruleSet') + "'"
    } elseif ($ln -eq 'exitRuleSet') {
        return "(stop processing further rules)"
    }
    return "Action: $ln"
}

function Get-Logic {
    # The centrepiece: one readable row per rule, with trigger + condition + actions in plain English.
    param($Root, $Nsm, $RsTriggers)
    $rows = @()
    $referencedRuleSets = @{}
    foreach ($k in $RsTriggers.Keys) { $referencedRuleSets[$k] = $true }
    # Also count ruleSets referenced by ruleSetAction inside rules (chained).
    foreach ($r in (Sel $Root '//xsf:rule/xsf:ruleSetAction' $Nsm)) {
        $rs = Get-Attr $r 'ruleSet'; if ($rs) { $referencedRuleSets[$rs] = $true }
    }

    foreach ($rs in (Sel $Root '//xsf:ruleSets/xsf:ruleSet' $Nsm)) {
        $rsName = Get-Attr $rs 'name'
        $trigger = 'Unknown (button or unused)'
        if ($RsTriggers.ContainsKey($rsName)) { $trigger = $RsTriggers[$rsName] }
        foreach ($rule in (Sel $rs './xsf:rule' $Nsm)) {
            $caption = Get-Attr $rule 'caption'
            $enabled = (Get-Attr $rule 'isEnabled') -ne 'no'
            $cond = To-PlainExpr (Get-Attr $rule 'condition')
            $actions = @()
            foreach ($act in $rule.ChildNodes) {
                if ($act.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                $actions += (Describe-Action $act)
            }
            $actionText = ($actions -join '; ')
            if (-not $actionText) { $actionText = '(no actions)' }

            # Find which views this rule might "appear" on based on its trigger field or button.
            $viewsForTrigger = ''
            if ($trigger -match '\[(?<f>[^\]]+)\]') {
                $tf = $Matches['f']
                if ($fieldUsage -and $fieldUsage.ContainsKey($tf)) { $viewsForTrigger = ($fieldUsage[$tf] | Select-Object -Unique) -join ', ' }
            } elseif ($trigger -match "on view '(?<v>[^']+)'") {
                $viewsForTrigger = $Matches['v']
            }

            $rows += [pscustomobject]@{
                Trigger = $trigger
                Views = $viewsForTrigger
                Rule = $caption
                Condition = $(if ($cond) { 'IF ' + $cond } else { '(always)' })
                Actions = $actionText
                Enabled = $enabled
                RuleSet = $rsName
            }
        }
    }
    return [pscustomobject]@{ Rows = $rows; ReferencedRuleSets = $referencedRuleSets }
}

function Get-OnLoadSummary {
    param($Root, $Nsm, $Adapters)
    $items = @()
    # rules that run on load
    $olSets = @()
    foreach ($r in (Sel $Root '//xsf:onLoad/xsf:ruleSetAction' $Nsm)) {
        $rs = Get-Attr $r 'ruleSet'; if ($rs) { $olSets += $rs }
    }
    if ($olSets.Count -gt 0) { $items += "Runs rule set(s): " + ($olSets -join ', ') }
    # queries on load
    $onLoadQueries = @($Adapters | Where-Object { $_.OnLoad })
    foreach ($q in $onLoadQueries) { $items += "Queries '" + $q.DataObject + "' (" + $q.Kind + ")" }
    return $items
}

function Get-EventHandlers {
    param($Root, $Nsm)
    $rows = @()
    foreach ($h in (Sel $Root '//xsf:domEventHandlers/xsf:domEventHandler' $Nsm)) {
        $leaf = Leaf-Name (Get-Attr $h 'match')
        $sets = @()
        foreach ($r in (Sel $h './xsf:ruleSetAction' $Nsm)) { $sets += (Get-Attr $r 'ruleSet') }
        $rows += [pscustomobject]@{ Field = $leaf; RuleSets = ($sets -join ', ') }
    }
    return $rows
}

function Get-Calculations {
    param($Root, $Nsm)
    $rows = @()
    foreach ($c in (Sel $Root '//xsf:calculatedField' $Nsm)) {
        $rows += [pscustomobject]@{
            Field = Leaf-Name (Get-Attr $c 'target')
            Formula = To-PlainExpr (Get-Attr $c 'expression')
            Recalc = Get-Attr $c 'refresh'
        }
    }
    return $rows
}

function Get-Validation {
    param($Root, $Nsm)
    $rows = @()
    foreach ($e in (Sel $Root '//xsf:errorCondition' $Nsm)) {
        $field = Leaf-Name (Get-Attr $e 'expressionContext')
        if (-not $field) { $field = Leaf-Name (Get-Attr $e 'match') }
        $msg = ''
        $m = Sel1 $e './xsf:errorMessage' $Nsm
        if ($m) { $msg = Get-Attr $m 'shortMessage' }
        $caption = ''
        $c = Sel1 $e './processing-instruction("caption")' $Nsm
        if ($c) { $caption = $c.Data.Trim() }
        $rows += [pscustomobject]@{
            Field = $field
            Rule = $caption
            FailsWhen = To-PlainExpr (Get-Attr $e 'expression')
            Message = $msg
        }
    }
    return $rows
}

function Get-Views {
    param($Root, $Nsm)
    $rows = @()
    $default = ''
    $vs = Sel1 $Root '//xsf:views' $Nsm
    if ($vs) { $default = Get-Attr $vs 'default' }
    # Views flagged client-only (do not render in the browser) live in the xsf3 extension.
    $clientOnly = @{}
    foreach ($ve in (Sel $Root '//xsf3:viewsExtension/xsf3:viewExtension' $Nsm)) {
        if ((Get-Attr $ve 'clientOnly') -eq 'yes') { $clientOnly[(Get-Attr $ve 'ref')] = $true }
    }
    foreach ($v in (Sel $Root '//xsf:views/xsf:view' $Nsm)) {
        $name = Get-Attr $v 'name'
        $mp = Sel1 $v './xsf:mainpane' $Nsm
        $xsl = ''
        if ($mp) { $xsl = Get-Attr $mp 'transform' }
        $readOnly = (Get-Attr $v 'designMode') -eq 'protected'
        $rows += [pscustomobject]@{
            View = $name; File = $xsl; IsDefault = ($name -eq $default); Caption = Get-Attr $v 'caption'
            ReadOnly = $readOnly; PrintView = Get-Attr $v 'printView'; ClientOnly = [bool]$clientOnly[$name]
        }
    }
    return $rows
}

function Get-Roles {
    # InfoPath user roles (xsf:roles) plus any identity/role references in expressions.
    param($Root, $Nsm, [string]$ManifestText)
    $roles = @()
    $rolesNode = Sel1 $Root '//xsf:roles' $Nsm
    if ($rolesNode) {
        $default = Get-Attr $rolesNode 'default'
        $initiator = Get-Attr $rolesNode 'initiator'
        foreach ($r in (Sel $rolesNode './xsf:role' $Nsm)) {
            $nm = Get-Attr $r 'name'
            $note = @()
            if ($nm -eq $default) { $note += 'default' }
            if ($nm -eq $initiator) { $note += 'initiator' }
            $members = @()
            foreach ($m in (Sel $Root "//xsf:membership/*[@memberOf='$nm']" $Nsm)) {
                $members += (Get-Attr $m 'name')
            }
            $roles += [pscustomobject]@{ Role = $nm; Notes = ($note -join ', '); Members = ($members -join '; ') }
        }
    }
    # Identity-driven logic (even without formal roles): functions that read the current user.
    $identitySignals = @()
    if ($ManifestText -match 'get-UserName') { $identitySignals += 'Uses current-user name (xdUser:get-UserName)' }
    if ($ManifestText -match 'isUserMemberOf|userRole') { $identitySignals += 'Branches on user role / group membership' }
    if ($ManifestText -match 'userprofileservice') { $identitySignals += 'Calls the SharePoint user-profile service' }
    return [pscustomobject]@{ Roles = $roles; IdentitySignals = $identitySignals }
}

function Get-Blockers {
    # Features that have no clean migration path - surfaced so a human can plan the rebuild.
    param($Root, $Nsm, [string]$FolderDir, [string]$ManifestText, $Adapters)
    $rows = @()
    # Managed code / script
    if ((Sel $Root '//xsf2:managedCode' $Nsm).Count -gt 0 -or (Get-ChildItem -LiteralPath $FolderDir -Filter '*.dll' -File -ErrorAction SilentlyContinue)) {
        $rows += [pscustomobject]@{ Blocker = 'Managed code-behind (.dll)'; Impact = 'High'; Note = 'Custom C#/VB logic must be rebuilt by hand.' }
    }
    if ((Sel $Root '//xsf:scripts/xsf:script' $Nsm).Count -gt 0) {
        $rows += [pscustomobject]@{ Blocker = 'Form script (JScript/VBScript)'; Impact = 'High'; Note = 'Scripted logic must be rebuilt.' }
    }
    # ActiveX controls
    if ((Sel $Root '//xsf:permissions/xsf:allowedControl' $Nsm).Count -gt 0 -or $ManifestText -match 'requireActiveX') {
        $rows += [pscustomobject]@{ Blocker = 'ActiveX control(s)'; Impact = 'Medium'; Note = 'e.g. people picker / custom ActiveX - no direct equivalent.' }
    }
    # Digital signatures
    if ((Sel $Root '//xsf:documentSignatures' $Nsm).Count -gt 0 -or (Sel $Root '//xsf3:signatureLine' $Nsm).Count -gt 0) {
        $rows += [pscustomobject]@{ Blocker = 'Digital signatures'; Impact = 'High'; Note = 'No equivalent in Power Apps / Lists forms.' }
    }
    # BCS
    if (@($Adapters | Where-Object { $_.Kind -like 'Business Connectivity*' }).Count -gt 0) {
        $rows += [pscustomobject]@{ Blocker = 'BCS / external content type'; Impact = 'High'; Note = 'External data source needs redesign.' }
    }
    # Direct SQL
    if (@($Adapters | Where-Object { $_.Kind -eq 'SQL / database' }).Count -gt 0) {
        $rows += [pscustomobject]@{ Blocker = 'Direct SQL/database connection'; Impact = 'High'; Note = 'Needs an on-prem data gateway or redesign.' }
    }
    # UDC data-connection-file references
    if ($ManifestText -match 'connectoid|\.udcx') {
        $rows += [pscustomobject]@{ Blocker = 'UDC data-connection file (.udcx)'; Impact = 'Medium'; Note = 'Connection lives in a data-connection library.' }
    }
    # REST / external web service
    if (@($Adapters | Where-Object { $_.Kind -eq 'Web service (SOAP)' }).Count -gt 0) {
        $rows += [pscustomobject]@{ Blocker = 'SOAP web-service connection'; Impact = 'Medium'; Note = 'Re-implement via Power Automate / connector.' }
    }
    return $rows
}

function Get-Navigation {
    # Button/event -> target view, derived from switchViewAction rules.
    param($Root, $Nsm, $RsTriggers)
    $rows = @()
    foreach ($rs in (Sel $Root '//xsf:ruleSets/xsf:ruleSet' $Nsm)) {
        $rsName = Get-Attr $rs 'name'
        $trigger = 'Unknown trigger'
        if ($RsTriggers.ContainsKey($rsName)) { $trigger = $RsTriggers[$rsName] }
        foreach ($sv in (Sel $rs './/xsf:switchViewAction' $Nsm)) {
            $rows += [pscustomobject]@{ Trigger = $trigger; GoesToView = (Get-Attr $sv 'view'); RuleSet = $rsName }
        }
    }
    return $rows
}

function Build-ViewFlowMermaid {
    # A Mermaid flowchart of view-to-view transitions. Best-effort: edges come from switchView
    # rules; the source view is approximated by the rule's trigger when it is a field on a view.
    param($Views, $Navigation, [string]$DefaultView)
    $lines = @()
    $lines += '```mermaid'
    $lines += 'flowchart TD'
    if (@($Views).Count -eq 0) { $lines += '  none[No views found]'; $lines += '```'; return ($lines -join "`n") }
    $idOf = @{}
    $i = 0
    foreach ($v in $Views) {
        $i++
        $id = "V$i"
        $idOf[$v.View] = $id
        $label = $v.View
        if ($v.IsDefault) { $label = $label + ' (default)' }
        if ($v.ReadOnly) { $label = $label + ' [read-only]' }
        $lines += "  $id[""$label""]"
    }
    if (@($Navigation).Count -eq 0) {
        $lines += "  %% No switch-view navigation found between views."
    } else {
        foreach ($n in $Navigation) {
            $toId = $idOf[$n.GoesToView]
            if (-not $toId) { continue }
            # crude source resolution: default view as the hub if trigger is load/submit/unknown
            $fromId = $idOf[$DefaultView]
            if (-not $fromId -and $idOf.Count -gt 0) { $fromId = ($idOf.Values | Select-Object -First 1) }
            $lbl = ($n.Trigger -replace '"', "'")
            if ($fromId -and $toId) { $lines += "  $fromId -->|""$lbl""| $toId" }
        }
    }
    $lines += '```'
    return ($lines -join "`n")
}

function Get-CleanLogicXml {
    # A compact, LLM-friendly extract of ONLY the logic-bearing manifest elements, with all
    # xmlns declarations and inter-tag whitespace stripped. This keeps the exact expressions /
    # field paths / adapter GUIDs an LLM needs to rebuild the form, at a fraction of the tokens
    # of the full manifest (no namespace soup, no packaged-file list, no design-mode chrome,
    # no schema/extension boilerplate).
    param($Root, $Nsm)
    $parts = @()
    foreach ($w in @('views', 'submit', 'dataObjects', 'ruleSets', 'domEventHandlers', 'onLoad', 'calculations', 'customValidation')) {
        $n = Sel1 $Root "//xsf:$w" $Nsm
        if ($n -and $n.OuterXml) { $parts += $n.OuterXml }
    }
    # The form's own list/columns (top-level query adapter) + any standalone validation.
    $q = Sel1 $Root '/xsf:xDocumentClass/xsf:query' $Nsm
    if ($q -and $q.OuterXml) { $parts += $q.OuterXml }
    foreach ($ec in (Sel $Root '//xsf:customValidation/xsf:errorCondition' $Nsm)) {
        # already inside customValidation above; skip duplicates
    }
    $xml = ($parts -join "`n")
    $xml = $xml -replace '(?s)<xsf:editing>.*?</xsf:editing>', ''   # per-field UI binding noise (in Fields sheet)
    $xml = $xml -replace '\s+xmlns(:\w+)?="[^"]*"', ''              # drop namespace declarations
    $xml = $xml -replace '>\s+<', '><'                              # collapse inter-tag whitespace
    # SECURITY: never emit plaintext SQL credentials (the PO forms embed Password=... in a
    # connectionString). Redact password / pwd anywhere they appear.
    $xml = $xml -replace '(?i)(password|pwd)\s*=\s*[^;"&]*', '$1=***REDACTED***'
    return $xml
}

function Clean-XslText {
    # A much smaller, readable version of a view .xsl: drop namespace declarations, CSS / cosmetic
    # attributes and whitespace, but KEEP the table structure, control bindings, labels, and the
    # xsl:if/choose conditions - i.e. the exact layout, minus the noise.
    param([string]$Txt)
    if (-not $Txt) { return '' }
    $Txt = $Txt -replace '(?s)<!--.*?-->', ''
    $Txt = $Txt -replace '(?s)<xsl:attribute name="style">.*?</xsl:attribute>', ''
    $Txt = $Txt -replace '\s+(style|class|tabIndex|hideFocus|contentEditable|xd:CtrlId|xd:boundProp|align|valign|width|height|cellpadding|cellspacing|border|bgcolor|nowrap|hideFocus)="[^"]*"', ''
    $Txt = $Txt -replace '\s+xmlns(:\w+)?="[^"]*"', ''
    $Txt = $Txt -replace '>\s+<', '><'
    return $Txt
}

function To-MdTable {
    # Render an array of PSCustomObjects as a GitHub markdown table (for the LLM context file).
    param($Rows, [string]$Title)
    $out = @()
    if ($Title) { $out += "### $Title"; $out += '' }
    $arr = @($Rows)
    if ($arr.Count -eq 0) { $out += '_None._'; $out += ''; return ($out -join "`n") }
    $cols = $arr[0].PSObject.Properties.Name
    $out += '| ' + ($cols -join ' | ') + ' |'
    $out += '| ' + (($cols | ForEach-Object { '---' }) -join ' | ') + ' |'
    foreach ($r in $arr) {
        $cells = @()
        foreach ($c in $cols) {
            $val = [string]$r.$c
            $val = $val -replace '\r?\n', ' ' -replace '\|', '\'
            $cells += $val
        }
        $out += '| ' + ($cells -join ' | ') + ' |'
    }
    $out += ''
    return ($out -join "`n")
}

# ============================================================================================
#  XSL (view) analysis - regex scrape, namespace-tolerant
# ============================================================================================

function Get-LeadingLabel {
    # The user-visible caption for a control: the last visible text immediately before the control's
    # opening tag (InfoPath lays labels out in the preceding table cell). Best-effort but reliable
    # in practice (e.g. my:Phone_Type -> "Requested Phone Type:").
    param([string]$Text, [int]$Pos)
    $start = [Math]::Max(0, $Pos - 600)
    $before = $Text.Substring($start, $Pos - $start)
    $parts = @(($before -replace '<[^>]+>', '|') -split '\|' | ForEach-Object { (Decode-Xml ($_ -replace '\s+', ' ')).Trim() } | Where-Object { $_ })
    if ($parts.Count -eq 0) { return '' }
    # Reject leaked CSS / markup fragments / control-state words that are not real labels.
    $reject = '(?i)(display\s*:|background|font-|width\s*:|height\s*:|margin|padding|border|color\s*:|#[0-9a-f]{3,6}|\d\s*p[xt]\b|align=|caption\s*:|\bchecked\b|\bselected\b|normal"|[<>{}"]|^(true|false)$)'
    # Prefer the last clean part that ends with ':' (InfoPath labels almost always do).
    for ($i = $parts.Count - 1; $i -ge 0; $i--) {
        $p = $parts[$i]
        if ($p.EndsWith(':') -and $p -match '[A-Za-z]' -and $p -notmatch $reject -and $p.Length -le 70) { return $p }
    }
    # Otherwise the last clean, substantial part (>=5 chars or multi-word, to skip stray fragments).
    for ($i = $parts.Count - 1; $i -ge 0; $i--) {
        $p = $parts[$i]
        if ($p -match '[A-Za-z]' -and $p -notmatch $reject -and $p.Length -le 70 -and ($p.Length -ge 5 -or $p -match '\s')) { return $p }
    }
    return ''
}

function Analyze-Xsl {
    param([string]$Path)
    $txt = ''
    try { $txt = [System.IO.File]::ReadAllText($Path) } catch { return $null }

    # Controls + bindings
    $controls = @()
    $readonly = @()    # field -> condition under which it is read-only/disabled ('' = always)
    $bindings = New-Object System.Collections.Generic.HashSet[string]
    foreach ($m in [regex]::Matches($txt, '<[^>]*xd:xctname="(?<ct>[^"]*)"[^>]*>')) {
        $tag = $m.Value
        $ct = $m.Groups['ct'].Value
        if ($ct -match '61e40d31') { $ct = 'PeoplePicker' }
        $bind = ''
        $bm = [regex]::Match($tag, 'xd:binding="(?<b>[^"]*)"')
        if ($bm.Success) { $bind = Leaf-Name $bm.Groups['b'].Value; if ($bind) { [void]$bindings.Add($bind) } }
        $label = ''
        if ($bind) { $label = Get-LeadingLabel $txt $m.Index }
        # Unconditional read-only.
        if ($bind -and $tag -match 'xd:disableEditing="yes"') { $readonly += [pscustomobject]@{ Field = $bind; Condition = ''; Caption = '' } }
        $controls += [pscustomobject]@{ Control = $ct; Field = $bind; Label = $label }

        # Capture button-triggered ruleSets
        $rsAction = [regex]::Match($tag, 'xd:xdxsfAction="(?<rs>[^"]*)"')
        if ($rsAction.Success) {
            $rsName = $rsAction.Groups['rs'].Value
            $controls += [pscustomobject]@{ Control = "Rule Trigger ($rsName)"; Field = $bind; Label = $label }
        }
    }

    # Formatting rule captions (human-readable names) from 'style' attributes.
    $ruleCaptions = @{}
    foreach ($sm in [regex]::Matches($txt, '(?s)<xsl:attribute name="style">.*?</xsl:attribute>')) {
        $before = $txt.Substring(0, $sm.Index)
        $bm = [regex]::Matches($before, 'xd:binding="(?<b>[^"]*)"')
        if ($bm.Count -eq 0) { continue }
        $fld = Leaf-Name $bm[$bm.Count - 1].Groups['b'].Value
        foreach ($m in [regex]::Matches($sm.Value, 'test="(?<c>[^"]+)"[^>]*>caption:\s*(?<cap>[^<]+)</xsl:')) {
            $cond = $m.Groups['c'].Value
            if (-not $ruleCaptions[$fld]) { $ruleCaptions[$fld] = @{} }
            $ruleCaptions[$fld][$cond] = $m.Groups['cap'].Value.Trim()
        }
    }

    # Conditional read-only / disabled: <xsl:attribute name="disabled"> or "contentEditable",
    # attributed to the nearest preceding bound control.
    foreach ($m in [regex]::Matches($txt, '(?s)<xsl:attribute name="(disabled|contentEditable)">.*?</xsl:attribute>')) {
        $attrName = $m.Groups[1].Value
        $attrBody = $m.Value
        $before = $txt.Substring(0, $m.Index)
        $bm = [regex]::Matches($before, 'xd:binding="(?<b>[^"]*)"')
        if ($bm.Count -eq 0) { continue }
        $fld = Leaf-Name $bm[$bm.Count - 1].Groups['b'].Value
        # disabled="true" or contentEditable="false"
        $isReadOnly = ($attrName -eq 'disabled' -and $attrBody -match '>true<') -or ($attrName -eq 'contentEditable' -and $attrBody -match '>false<')
        if (-not $isReadOnly) { continue }

        $cond = ""; $tm = [regex]::Match($attrBody, 'test="(?<c>[^"]+)"')
        if ($tm.Success) { $cond = $tm.Groups['c'].Value }
        else {
            # Case B: condition is in a parent xsl:if or xsl:when
            $around = $txt.Substring([Math]::Max(0, $m.Index - 500), 500)
            $tm2 = [regex]::Matches($around, 'test="(?<c>[^"]+)"'); if ($tm2.Count -gt 0) { $cond = $tm2[$tm2.Count - 1].Groups['c'].Value }
        }
        if ($cond) {
            $cap = ""; if ($ruleCaptions[$fld] -and $ruleCaptions[$fld][$cond]) { $cap = $ruleCaptions[$fld][$cond] }
            $readonly += [pscustomobject]@{ Field = $fld; Condition = $cond; Caption = $cap }
        }
    }

    # Sections / repeating
    $sectionCount = ([regex]::Matches($txt, 'xd:xctname="(Section|RepeatingSection|RepeatingTable)"')).Count

    # Conditional visibility: <xsl:if test="EXPR"> ... and detect a nearby display:none (hide-when).
    $vis = @()
    foreach ($m in [regex]::Matches($txt, '<xsl:if\s+test="(?<t>[^"]{3,300})">')) {
        $test = $m.Groups['t'].Value
        if ($test -match 'function-available') { continue }
        $window = $txt.Substring($m.Index, [Math]::Min(400, $txt.Length - $m.Index))
        $hideWhen = ($window -match '(?i)display:\s*none')
        $vis += [pscustomobject]@{ TestRaw = $test; HideWhen = $hideWhen }
    }

    # Dropdown / option controls
    $dropdowns = @()
    foreach ($sm in [regex]::Matches($txt, '(?s)<select[^>]*xd:xctname="(?<ct>dropdown|combobox|multiselectlistbox|ListBox)"[^>]*>(?<body>.*?)</select>')) {
        $tag = $sm.Value
        $bind = ''
        $bm = [regex]::Match($tag, 'xd:binding="(?<b>[^"]*)"')
        if ($bm.Success) { $bind = Leaf-Name $bm.Groups['b'].Value }
        $body = $sm.Groups['body'].Value
        $opts = @()
        foreach ($om in [regex]::Matches($body, '(?s)<option\s+value="(?<v>[^"]*)"[^>]*>(?<txt>.*?)</option>')) {
            $v = (Decode-Xml $om.Groups['v'].Value).Trim()
            # The option body wraps the display label in xsl:if/xsl:attribute that sets 'selected'.
            # Drop those subtrees first (their text 'selected' is not part of the label), then strip
            # remaining tags and collapse whitespace to get the clean display text.
            $inner = $om.Groups['txt'].Value -replace '(?s)<xsl:attribute\b[^>]*>.*?</xsl:attribute>', ''
            $t = (Decode-Xml (($inner -replace '<[^>]+>', ' ') -replace '\s+', ' ')).Trim()
            $label = ''
            if ($t -and $v -and $t -ne $v) { $label = "$t ($v)" }   # show both when display != stored value
            elseif ($t) { $label = $t }
            elseif ($v) { $label = $v }
            if ($label -ne '') { $opts += $label }
        }
        $opts = @($opts | Select-Object -Unique)
        $dynamic = ''
        $dm = [regex]::Match($body, "GetDOM\('(?<c>[^']*)'\)")
        if ($dm.Success) { $dynamic = $dm.Groups['c'].Value }
        $dropdowns += [pscustomobject]@{
            Field = $bind
            Type = $(if ($dynamic) { 'Dynamic (from ' + $dynamic + ')' } else { 'Static' })
            Options = $(if ($opts.Count -gt 0) { ($opts | Select-Object -First 25) -join ' | ' } else { '' })
        }
    }

    return [pscustomobject]@{
        Controls = $controls; Bindings = $bindings; SectionCount = $sectionCount
        Visibility = $vis; Dropdowns = $dropdowns; Readonly = $readonly
    }
}

# ============================================================================================
#  File classification + orphan move
# ============================================================================================

function Add-SchemaClosure {
    param([string]$FolderDir, [string]$SchemaFile, $Active, $ActiveReason, [string]$Reason)
    if (-not $SchemaFile) { return }
    if ($Active.Contains($SchemaFile)) { return }
    [void]$Active.Add($SchemaFile)
    if ($ActiveReason -and -not $ActiveReason.ContainsKey($SchemaFile)) { $ActiveReason[$SchemaFile] = $Reason }
    $path = Join-Path $FolderDir $SchemaFile
    if (-not (Test-Path -LiteralPath $path)) { return }
    $txt = ''
    try { $txt = [System.IO.File]::ReadAllText($path) } catch { return }
    foreach ($m in [regex]::Matches($txt, 'schemaLocation="(?<f>[^"]+)"')) {
        # Imported partner schemas (the InfoPath dataFields/queryFields/PartnerControls/
        # documentManagement namespace companions of the schema above).
        Add-SchemaClosure -FolderDir $FolderDir -SchemaFile $m.Groups['f'].Value -Active $Active -ActiveReason $ActiveReason -Reason 'Partner schema (imported by another schema)'
    }
}

function Get-OrphanReason {
    param([string]$Name, [string]$Ext, $UpgradeTransforms)
    if ($UpgradeTransforms -and $UpgradeTransforms.Contains($Name)) { return 'Legacy schema-upgrade transform (no business logic)' }
    if ($Name -imatch '\.dll\d+\.config$') { return 'Superseded code-config version' }
    if ($Ext -eq '.pdb') { return 'Debug symbols (not needed)' }
    if ($Name -ieq 'sampledata.xml') { return 'Design-time sample data' }
    if ($Ext -eq '.xsd') { return 'Superseded / unreferenced schema version' }
    if ($Ext -eq '.xsl') { return 'Legacy schema-upgrade transform (no business logic)' }
    return 'Not referenced by the form'
}

function Classify-And-MoveFiles {
    param([string]$FolderDir, $Root, $Nsm, $Views, $Adapters)
    $active = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $activeReason = @{}   # name -> specific reason (case-insensitive lookups)
    [void]$active.Add('manifest.xsf')

    # template / initial document
    $init = Sel1 $Root '//xsf:fileNew/xsf:initialXmlDocument' $Nsm
    if ($init) { $h = Get-Attr $init 'href'; if ($h) { [void]$active.Add($h); $activeReason[$h] = 'Form template (initial/default values)' } }
    # views
    foreach ($v in $Views) { if ($v.File) { [void]$active.Add($v.File); $activeReason[$v.File] = "View '$($v.View)'" } }
    # NOTE: upgrade transforms (xsf:documentVersionUpgrade/useTransform, i.e. upgrade.xsl) are
    # deliberately NOT added to the active set. They only reshape OLD saved instances to the
    # current schema and carry no business logic, so they are moved to _Unused.
    $upgradeTransforms = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($u in (Sel $Root '//xsf:documentVersionUpgrade//xsf:useTransform' $Nsm)) {
        [void]$upgradeTransforms.Add((Get-Attr $u 'transform'))
    }
    # schemas: dataObject @schema + documentSchema locations, each with xsd:import closure
    foreach ($d in (Sel $Root '//xsf:dataObjects/xsf:dataObject' $Nsm)) {
        Add-SchemaClosure -FolderDir $FolderDir -SchemaFile (Get-Attr $d 'schema') -Active $active -ActiveReason $activeReason -Reason "Schema for data connection '$(Get-Attr $d 'name')'"
    }
    foreach ($ds in (Sel $Root '//xsf:documentSchemas/xsf:documentSchema' $Nsm)) {
        $loc = Get-Attr $ds 'location'
        $f = $loc; if ($loc -match '\s') { $f = ($loc -split '\s+')[-1] }
        $rsn = $(if ((Get-Attr $ds 'rootSchema') -eq 'yes') { 'Main form schema (root)' } else { 'Main form schema (partner namespace)' })
        Add-SchemaClosure -FolderDir $FolderDir -SchemaFile $f -Active $active -ActiveReason $activeReason -Reason $rsn
    }
    # web service input files
    foreach ($inp in (Sel $Root '//xsf:webServiceAdapter//xsf:input' $Nsm)) {
        $src = Get-Attr $inp 'source'; if ($src) { [void]$active.Add($src); $activeReason[$src] = 'Web-service request template' }
    }
    # code-behind: the dll + base .dll.config (numbered versions are junk)
    $hasCode = $false
    $dllFiles = Get-ChildItem -LiteralPath $FolderDir -Filter '*.dll' -File -ErrorAction SilentlyContinue
    foreach ($dll in $dllFiles) {
        $hasCode = $true
        [void]$active.Add($dll.Name); $activeReason[$dll.Name] = 'Managed code-behind library'
        [void]$active.Add($dll.Name + '.config'); $activeReason[$dll.Name + '.config'] = 'Code-behind configuration'
    }
    # images referenced by any active xsl
    foreach ($v in $Views) {
        if (-not $v.File) { continue }
        $vp = Join-Path $FolderDir $v.File
        if (-not (Test-Path -LiteralPath $vp)) { continue }
        $vt = ''
        try { $vt = [System.IO.File]::ReadAllText($vp) } catch { $vt = '' }
        foreach ($m in [regex]::Matches($vt, '(?i)([A-Za-z0-9_\- ]+\.(png|gif|jpg|jpeg|bmp))')) {
            $img = $m.Groups[1].Value
            if (-not $active.Contains($img)) { [void]$active.Add($img); $activeReason[$img] = 'Image used in a view' }
        }
    }

    # Now classify every top-level file.
    $fileRows = @()
    $unusedDir = Join-Path $FolderDir '_Unused'
    $movedCount = 0
    $files = Get-ChildItem -LiteralPath $FolderDir -File -ErrorAction SilentlyContinue
    foreach ($f in $files) {
        $nm = $f.Name
        $ext = $f.Extension.ToLowerInvariant()
        if ($nm -ieq 'manifest.xsf') {
            $fileRows += [pscustomobject]@{ File = $nm; Status = 'Active'; Reason = 'Form definition' }; continue
        }
        # Never touch this tool's own output files (matters on re-runs).
        if ($ext -eq '.xlsx' -or $ext -eq '.txt' -or $ext -eq '.md') { continue }
        $isActive = $active.Contains($nm)
        $status = 'Active'; $reason = ''
        if ($isActive) {
            $reason = $activeReason[$nm]
            if (-not $reason) { $reason = 'Referenced by the form' }
        } else {
            $status = 'Orphaned'
            $reason = Get-OrphanReason -Name $nm -Ext $ext -UpgradeTransforms $upgradeTransforms
        }
        if ($status -eq 'Orphaned') {
            if (-not (Test-Path -LiteralPath $unusedDir)) { New-Item -ItemType Directory -Path $unusedDir -Force | Out-Null }
            $dest = Join-Path $unusedDir $nm
            try { Move-Item -LiteralPath $f.FullName -Destination $dest -Force; $movedCount++ }
            catch { $reason = $reason + " (move failed: $($_.Exception.Message))" }
        }
        $fileRows += [pscustomobject]@{ File = $nm; Status = $status; Reason = $reason }
    }
    # Also account for files already quarantined in _Unused on an earlier run, so the Files sheet
    # and the orphaned count are IDENTICAL whether this is a first run or a re-run.
    $seen = @{}
    foreach ($r in $fileRows) { $seen[$r.File] = $true }
    if (Test-Path -LiteralPath $unusedDir) {
        foreach ($uf in (Get-ChildItem -LiteralPath $unusedDir -File -ErrorAction SilentlyContinue)) {
            if ($seen[$uf.Name]) { continue }
            $reason = Get-OrphanReason -Name $uf.Name -Ext $uf.Extension.ToLowerInvariant() -UpgradeTransforms $upgradeTransforms
            $fileRows += [pscustomobject]@{ File = $uf.Name; Status = 'Orphaned'; Reason = $reason }
        }
    }
    $orphanTotal = @($fileRows | Where-Object { $_.Status -eq 'Orphaned' }).Count
    return [pscustomobject]@{ Rows = $fileRows; HasCodeBehind = $hasCode; MovedCount = $movedCount; OrphanedTotal = $orphanTotal }
}

# ============================================================================================
#  Excel writing
# ============================================================================================

function Remove-FileSafe {
    # OneDrive briefly locks a freshly-written .xlsx while it syncs, so a same-path overwrite on a
    # re-run can throw "the process cannot access the file." Retry a few times before giving up.
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $true }
    for ($i = 0; $i -lt 12; $i++) {
        try { Remove-Item -LiteralPath $Path -Force -ErrorAction Stop; return $true }
        catch { Start-Sleep -Milliseconds 1000 }
    }
    return $false
}

function Add-Sheet {
    # Adds one worksheet to a shared in-memory ExcelPackage and returns the package, so the whole
    # workbook is written to disk ONCE (at Close-ExcelPackage). This is faster than re-opening the
    # file per sheet and - crucially on OneDrive paths - shrinks the file-lock window to a single
    # write instead of ~14, so sync activity cannot corrupt a multi-sheet build.
    param($Package, [string]$Path, [string]$Name, $Rows, [switch]$NoTable, [switch]$NoAutoSize)
    if ($null -eq $Rows -or @($Rows).Count -eq 0) {
        $Rows = @([pscustomobject]@{ Note = 'None found' })
    }
    $params = @{
        WorksheetName = $Name; FreezeTopRow = $true; PassThru = $true
    }
    if (-not $NoAutoSize) { $params['AutoSize'] = $true }
    if ($Package) { $params['ExcelPackage'] = $Package } else { $params['Path'] = $Path }
    if (-not $NoTable) { $params['TableName'] = ($Name -replace '[^A-Za-z0-9]', ''); $params['TableStyle'] = 'Medium2' }
    return (@($Rows) | Export-Excel @params)
}

# ============================================================================================
#  Per-form driver
# ============================================================================================

function Analyze-Form {
    param([string]$XsnPath)

    $name = [System.IO.Path]::GetFileNameWithoutExtension($XsnPath)
    $folder = Join-Path $scriptRoot $name
    Write-Host ("  Extracting and analyzing: {0}" -f $name)

    if (-not (Expand-Xsn -XsnPath $XsnPath -DestDir $folder)) {
        Log-Issue $name 'extract' 'Skipped - could not extract the .xsn (form NOT in the report).' 'Error'
        return $null
    }

    $manifestPath = Join-Path $folder 'manifest.xsf'
    $xml = $null
    try { $xml = Load-Xml $manifestPath }
    catch { Log-Issue $name 'manifest' "Could not parse manifest.xsf (form NOT in the report): $($_.Exception.Message)" 'Error'; return $null }
    $root = $xml.DocumentElement
    $nsm = New-NsMgr $xml

    # ---- analyzers (each guarded so one failure does not abort the form) ----
    $ft = $null; try { $ft = Get-FormType $root $nsm } catch { Log-Issue $name 'FormType' $_.Exception.Message }
    $adapters = @(); try { $adapters = @(Get-Adapters $root $nsm) } catch { Log-Issue $name 'Adapters' $_.Exception.Message }
    $primary = $null; try { $primary = Get-PrimaryFields $root $nsm } catch { Log-Issue $name 'PrimaryFields' $_.Exception.Message }
    $views = @(); try { $views = @(Get-Views $root $nsm) } catch { Log-Issue $name 'Views' $_.Exception.Message }
    $rsTrig = @{}; try { $rsTrig = Build-RuleSetTriggers $root $nsm } catch { Log-Issue $name 'Triggers' $_.Exception.Message }
    $valid = @(); try { $valid = @(Get-Validation $root $nsm) } catch { Log-Issue $name 'Validation' $_.Exception.Message }
    $schemaFields = @(); try { $schemaFields = @(Get-SchemaFields $folder $root $nsm) } catch { Log-Issue $name 'SchemaFields' $_.Exception.Message }
    $structure = @(); try { $structure = @(Get-FormStructure $folder $root $nsm) } catch { Log-Issue $name 'Structure' $_.Exception.Message }
    $sectionCount = @($structure | Where-Object { $_.Kind -eq 'Section' -or $_.Kind -eq 'Repeating section' }).Count
    $repeatingSections = @($structure | Where-Object { $_.Kind -eq 'Repeating section' }).Count
    $repeatingFields = @($structure | Where-Object { $_.Kind -eq 'Repeating field' }).Count

    # ---- view (xsl) analysis ----
    $allBindings = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $visibilityRows = @()
    $dropdownRows = @()
    $readonlyRows = @()
    $buttonTriggers = @{}
    $sectionTotal = 0
    $fieldUsage = @{}        # field -> list of views
    $controlByField = @{}    # field -> friendly control type (rich text, date picker, attachment...)
    $labelByField = @{}      # field -> user-visible caption (e.g. "Requested Phone Type:")
    foreach ($v in $views) {
        if (-not $v.File) { continue }
        $vp = Join-Path $folder $v.File
        if (-not (Test-Path -LiteralPath $vp)) { continue }
        $xa = Analyze-Xsl $vp
        if ($null -eq $xa) { continue }
        $sectionTotal += $xa.SectionCount
        foreach ($c in $xa.Controls) {
            if ($c.Control -match 'Rule Trigger \((?<rs>.*)\)') {
                $rsName = $Matches['rs']
                if (-not $buttonTriggers.ContainsKey($rsName)) { $buttonTriggers[$rsName] = @() }
                $lbl = $c.Label; if (-not $lbl) { $lbl = "Button" }
                $buttonTriggers[$rsName] += "Click '$lbl' on view '$($v.View)'"
                continue
            }
            if (-not $c.Field) { continue }
            $fc = Friendly-Control $c.Control
            if ($fc -and $fc -ne 'Button' -and $fc -ne 'Section' -and -not $controlByField.ContainsKey($c.Field)) { $controlByField[$c.Field] = $fc }
            if ($c.Label -and -not $labelByField.ContainsKey($c.Field)) { $labelByField[$c.Field] = $c.Label }
        }
        foreach ($b in $xa.Bindings) {
            [void]$allBindings.Add($b)
            if (-not $fieldUsage.ContainsKey($b)) { $fieldUsage[$b] = @() }
            $fieldUsage[$b] += $v.View
        }
        foreach ($vis in $xa.Visibility) {
            $cond = To-PlainExpr $vis.TestRaw
            $pfx = To-PowerFx $vis.TestRaw
            if ($vis.HideWhen) { $pfx = "Not(" + $pfx + ")" }
            $visibilityRows += [pscustomobject]@{
                View = $v.View
                ShownWhen = $(if ($vis.HideWhen) { "Hidden when: $cond" } else { "Shown when: $cond" })
                DerivedPowerFxVisible = $pfx
            }
        }
        foreach ($dd in $xa.Dropdowns) {
            $dropdownRows += [pscustomobject]@{ View = $v.View; Field = $dd.Field; Type = $dd.Type; Options = $dd.Options }
        }
        foreach ($ro in $xa.Readonly) {
            $fld = Leaf-Name $ro.Field
            if ($ro.Condition) {
                $readonlyRows += [pscustomobject]@{
                    View = $v.View; Field = $fld; Rule = $ro.Caption
                    ReadOnlyWhen = "When " + (To-PlainExpr $ro.Condition)
                    DerivedDisplayMode = "If(" + (To-PowerFx $ro.Condition) + ", DisplayMode.View, DisplayMode.Edit)"
                }
            } else {
                $readonlyRows += [pscustomobject]@{ View = $v.View; Field = $fld; Rule = ''; ReadOnlyWhen = 'Always (read-only)'; DerivedDisplayMode = 'DisplayMode.View' }
            }
        }
    }
    $readonlyRows = @($readonlyRows | Sort-Object View, Field, ReadOnlyWhen -Unique)

    # Merge button triggers into ruleSet triggers
    foreach ($rs in $buttonTriggers.Keys) {
        $trig = ($buttonTriggers[$rs] | Select-Object -Unique) -join '; '
        if ($rsTrig.ContainsKey($rs)) { $rsTrig[$rs] += " (also: $trig)" }
        else { $rsTrig[$rs] = $trig }
    }

    # ---- logic analysis (requires fieldUsage from XSL analysis) ----
    $logic = $null; try { $logic = Get-Logic $root $nsm $rsTrig } catch { Log-Issue $name 'Logic' $_.Exception.Message }
    $onLoad = @(); try { $onLoad = @(Get-OnLoadSummary $root $nsm $adapters) } catch { Log-Issue $name 'OnLoad' $_.Exception.Message }
    $events = @(); try { $events = @(Get-EventHandlers $root $nsm) } catch { Log-Issue $name 'Events' $_.Exception.Message }
    $calcs = @(); try { $calcs = @(Get-Calculations $root $nsm) } catch { Log-Issue $name 'Calcs' $_.Exception.Message }

    # ---- merge fields (list columns + schema), mark usage / unused / logic-only ----
    # collect every field name referenced by any logic expression (rules/calcs/validation/conditions)
    $logicRefs = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $manifestText = ''
    try { $manifestText = [System.IO.File]::ReadAllText($manifestPath) } catch {}
    foreach ($m in [regex]::Matches($manifestText, '(?:my|d|q):(?<n>[A-Za-z0-9_]+)')) { [void]$logicRefs.Add($m.Groups['n'].Value) }

    # static default values from the template
    $defaults = @{}
    try { $defaults = Get-Defaults $folder $root $nsm } catch { Log-Issue $name 'Defaults' $_.Exception.Message }

    # which fields are SharePoint columns vs XML-only. List forms: every field is a column. Library
    # (XML) forms: only the property-promoted fields are columns; the rest live only in the XML.
    $promoted = @{}
    try { $promoted = Get-PromotedFields $root $nsm } catch { Log-Issue $name 'Promotion' $_.Exception.Message }
    $isListForm = ($ft -and $ft.Kind -eq 'SharePoint List form')

    # field -> parent section name (from the structure tree). Library forms have named sections;
    # list forms are flat, so this stays blank for them.
    $sectionByField = @{}
    $structRootName = ''
    if (@($structure).Count -gt 0) { $structRootName = $structure[0].Name }
    foreach ($s in $structure) {
        if ($s.Kind -like '*ection*') { continue }
        $segs = $s.Path -split '/'
        $section = ''
        if ($segs.Count -ge 2) {
            $parent = $segs[$segs.Count - 2]
            if ($parent -and $parent -ne $structRootName) { $section = $parent }
        }
        if ($section -and -not $sectionByField.ContainsKey($s.Name)) { $sectionByField[$s.Name] = $section }
    }

    # field -> human-readable validation-based requirements
    $validationReqs = @{}
    foreach ($v in $valid) {
        # 'cannot be blank' rules (requiredness)
        if ($v.FailsWhen -match 'is (blank|empty|""|"")' -or $v.FailsWhen -match '= ""' -or $v.Message -match 'Cannot be blank') {
            if (-not $validationReqs.ContainsKey($v.Field)) { $validationReqs[$v.Field] = @() }
            $label = $v.Rule; if (-not $label) { $label = 'Required' }
            # If condition mentions a view, tag it.
            if ($v.FailsWhen -match 'view-name is "(?<v>[^"]+)"') { $label += " (on view '$($Matches['v'])')" }
            $validationReqs[$v.Field] += $label
        }
    }

    $fieldRows = @()
    $fieldSeen = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $complexCount = 0
    $unusedCount = 0
    $baseFields = @()
    if ($primary -and $primary.Fields.Count -gt 0) { $baseFields = $primary.Fields }
    foreach ($f in $baseFields) {
        if ($script:StructuralWrappers -contains $f.Name) { continue }
        if (-not $fieldSeen.Add($f.Name)) { continue }
        $usedInViews = @()
        if ($fieldUsage.ContainsKey($f.Name)) { $usedInViews = $fieldUsage[$f.Name] | Select-Object -Unique }
        $inUi = $usedInViews.Count -gt 0
        $inLogic = $logicRefs.Contains($f.Name)
        $status = 'Active'
        if ($script:SystemFields -contains $f.Name) { $status = 'System' }
        elseif (-not $inUi -and -not $inLogic) { $status = 'Unused'; $unusedCount++ }
        elseif (-not $inUi -and $inLogic) { $status = 'Logic-only (no UI)' }
        if ($f.IsComplex) { $complexCount++ }
        $lt = ''
        if ($f.PSObject.Properties.Name -contains 'LookupTarget') { $lt = $f.LookupTarget }
        $storage = $(if ($isListForm) { 'SharePoint column' } elseif ($promoted.ContainsKey($f.Name)) { "SharePoint column ($($promoted[$f.Name]))" } else { 'XML only' })
        
        $req = $(if ($f.Required) { 'Yes (schema)' } else { '' })
        if ($validationReqs.ContainsKey($f.Name)) {
            $vreq = ($validationReqs[$f.Name] | Select-Object -Unique) -join '; '
            if ($req) { $req += "; " + $vreq } else { $req = "Conditional: " + $vreq }
        }
        if (-not $req) { $req = 'no' }

        $fieldRows += [pscustomobject]@{
            Section = $sectionByField[$f.Name]; Field = $f.Name; Label = $labelByField[$f.Name]; Type = $f.Type
            Control = $controlByField[$f.Name]; Storage = $storage; Complex = $f.IsComplex; LookupTarget = $lt
            Required = $req; Default = $defaults[$f.Name]; UsedInViews = ($usedInViews -join ', '); Status = $status
        }
    }
    # add schema-only structural fields (library forms / repeating groups) not already covered
    foreach ($f in $schemaFields) {
        if ($script:StructuralWrappers -contains $f.Name) { continue }
        if ($fieldSeen.Contains($f.Name)) { continue }
        # Skip pure container/group elements (no datatype and not a repeating group) - they are
        # structure, not data, and only add noise.
        if (-not $f.Type -and -not $f.Repeating) { continue }
        if (-not $fieldSeen.Add($f.Name)) { continue }
        $usedInViews = @()
        if ($fieldUsage.ContainsKey($f.Name)) { $usedInViews = $fieldUsage[$f.Name] | Select-Object -Unique }
        $storage = $(if ($isListForm) { 'SharePoint column' } elseif ($promoted.ContainsKey($f.Name)) { "SharePoint column ($($promoted[$f.Name]))" } else { 'XML only' })
        
        $req = $(if ($f.Nillable -eq $false) { 'Yes (schema)' } else { '' })
        if ($validationReqs.ContainsKey($f.Name)) {
            $vreq = ($validationReqs[$f.Name] | Select-Object -Unique) -join '; '
            if ($req) { $req += "; " + $vreq } else { $req = "Conditional: " + $vreq }
        }
        if (-not $req) { $req = 'no' }

        $fieldRows += [pscustomobject]@{
            Section = $sectionByField[$f.Name]; Field = $f.Name; Label = $labelByField[$f.Name]; Type = $f.Type
            Control = $controlByField[$f.Name]; Storage = $storage; Complex = $false; LookupTarget = ''
            Required = $req; Default = $defaults[$f.Name]; UsedInViews = ($usedInViews -join ', ')
            Status = $(if ($usedInViews.Count -gt 0) { 'Active' } elseif ($f.Repeating) { 'Repeating group' } else { 'Structure only' })
        }
    }

    # ---- connection counts ----
    $secondary = @($adapters | Where-Object { -not $_.IsPrimary })
    $spLists = @($adapters | Where-Object { $_.IsSharePointList })
    $primaryListId = ''
    if ($primary) { $primaryListId = $primary.PrimaryListId }
    $otherListIds = @($spLists | Where-Object { $_.ListId -and $_.ListId -ne $primaryListId } | Select-Object -ExpandProperty ListId -Unique)
    $sqlConns = @($adapters | Where-Object { $_.Kind -eq 'SQL / database' })
    $wsConns = @($adapters | Where-Object { $_.Kind -eq 'Web service (SOAP)' })
    $attachmentCount = @($fieldRows | Where-Object { $_.Control -eq 'File attachment' }).Count
    $spBoundCount = @($fieldRows | Where-Object { $_.Storage -like 'SharePoint*' }).Count
    $xmlOnlyCount = @($fieldRows | Where-Object { $_.Storage -eq 'XML only' }).Count

    # ---- classify + move files ----
    $files = $null
    try { $files = Classify-And-MoveFiles -FolderDir $folder -Root $root -Nsm $nsm -Views $views -Adapters $adapters }
    catch { Log-Issue $name 'Files' $_.Exception.Message; $files = [pscustomobject]@{ Rows = @(); HasCodeBehind = $false; MovedCount = 0; OrphanedTotal = 0 } }

    # ---- roles / identity, navigation, migration blockers ----
    $roles = $null
    try { $roles = Get-Roles $root $nsm $manifestText } catch { Log-Issue $name 'Roles' $_.Exception.Message; $roles = [pscustomobject]@{ Roles = @(); IdentitySignals = @() } }
    $navigation = @()
    try { $navigation = @(Get-Navigation $root $nsm $rsTrig) } catch { Log-Issue $name 'Navigation' $_.Exception.Message }
    $blockers = @()
    try { $blockers = @(Get-Blockers $root $nsm $folder $manifestText $adapters) } catch { Log-Issue $name 'Blockers' $_.Exception.Message }
    # More than one file-attachment control is a migration headache (each needs its own column/handling).
    if ($attachmentCount -gt 1) {
        $blockers += [pscustomobject]@{ Blocker = "Multiple file-attachment controls ($attachmentCount)"; Impact = 'Medium'; Note = 'Each attachment control needs its own column / handling in the target platform.' }
    }
    if ($repeatingSections -gt 0) {
        $blockers += [pscustomobject]@{ Blocker = "$repeatingSections repeating section(s)/table(s)"; Impact = 'Medium'; Note = 'Repeating data maps to a child list / gallery / collection - adds rebuild effort.' }
    }
    # Browser vs rich-client
    $runtime = ''
    $sd = Sel1 $root '//xsf2:solutionDefinition' $nsm
    if ($sd) { $runtime = Get-Attr $sd 'runtimeCompatibility' }
    $browserEnabled = ($runtime -match 'server')

    # ---- dead-logic / simplification findings ----
    $simplify = @()
    foreach ($lr in $logic.Rows) {
        if (-not $lr.Enabled) { $simplify += [pscustomobject]@{ Type = 'Disabled rule'; Item = $lr.Rule; Note = 'Rule exists but is turned off' } }
        elseif ($lr.Actions -eq '(no actions)') { $simplify += [pscustomobject]@{ Type = 'Empty rule'; Item = "$($lr.RuleSet): $($lr.Rule)"; Note = 'Rule has no actions - safe to remove' } }
        elseif ($lr.Trigger -like 'Unknown*') { $simplify += [pscustomobject]@{ Type = 'Possibly unused rule set'; Item = "$($lr.RuleSet): $($lr.Rule)"; Note = 'Rule set not wired to load or a field change' } }
    }
    foreach ($fr in $fieldRows) {
        if ($fr.Status -eq 'Unused') { $simplify += [pscustomobject]@{ Type = 'Unused field'; Item = $fr.Field; Note = 'Not shown in any view and not referenced by logic' } }
    }

    # ---- complexity score ----
    $rulesCount = @($logic.Rows).Count
    $condCount = @($logic.Rows | Where-Object { $_.Condition -ne '(always)' }).Count
    $navCount = @($navigation).Count
    $scoreParts = [ordered]@{
        Views               = @($views).Count * 2
        SecondaryConns      = $secondary.Count * 3
        OtherListConns      = $otherListIds.Count * 2
        SqlConns            = $sqlConns.Count * 5
        WebServiceConns     = $wsConns.Count * 3
        Fields              = [int]([math]::Round(@($fieldRows).Count * 0.5))
        ComplexFields       = $complexCount * 2
        Rules               = $rulesCount * 2
        RuleConditions      = $condCount * 1
        FieldChangeHandlers = @($events).Count * 2
        OnLoadActions       = @($onLoad).Count * 3
        Calculations        = @($calcs).Count * 1
        ValidationRules     = @($valid).Count * 1
        NavigationRules     = $navCount * 2
        ReadOnlyRules       = @($readonlyRows).Count * 1
        Sections            = [int]([math]::Round($sectionTotal * 0.5))
        RepeatingSections   = $repeatingSections * 6
        RepeatingFields     = $repeatingFields * 2
        Attachments         = [Math]::Max(0, $attachmentCount - 1) * 3
        XmlOnlyFields       = $xmlOnlyCount * 1
        CodeBehind          = $(if ($files.HasCodeBehind) { 40 } else { 0 })
    }
    $score = 0
    foreach ($k in $scoreParts.Keys) { $score += $scoreParts[$k] }
    $bucket = 'Low'
    if ($score -ge 150) { $bucket = 'Very High' }
    elseif ($score -ge 75) { $bucket = 'High' }
    elseif ($score -ge 30) { $bucket = 'Medium' }

    # ---- migration flags (plain English highlights) ----
    $flags = @()
    foreach ($b in $blockers) { $flags += ("[" + $b.Impact + "] " + $b.Blocker + " - " + $b.Note) }
    if ($otherListIds.Count -gt 0) { $flags += "$($otherListIds.Count) connection(s) to OTHER SharePoint lists." }
    if (@($views).Count -gt 1) { $flags += "$(@($views).Count) views (multi-screen form)." }
    if ($unusedCount -gt 0) { $flags += "$unusedCount custom field(s) appear unused - candidates to drop." }
    if (@($simplify | Where-Object { $_.Type -like 'Disabled*' }).Count -gt 0) { $flags += 'Contains disabled rules (dead logic).' }
    if (@($roles.Roles).Count -gt 0) { $flags += "$(@($roles.Roles).Count) InfoPath user role(s) defined - role-based behavior to reproduce." }
    foreach ($sig in $roles.IdentitySignals) { $flags += $sig }
    if (-not $browserEnabled -and $runtime) { $flags += 'Rich-client-only form (not browser-enabled).' }

    # ============================================================================
    #  Write the workbook
    # ============================================================================
    $xlsxFinal = Join-Path $folder ("{0}-Analysis.xlsx" -f $name)
    $xlsx = Join-Path $folder ("~{0}.build.xlsx" -f $name)   # build here, then atomically replace
    [void](Remove-FileSafe $xlsx)
    $pkg = $null

    # Per-form data quality: issues logged for THIS form by the analyzers that ran above.
    $formIssues = @($script:RunLog | Where-Object { $_.Form -eq $name }).Count

    # 1) Summary - a short identity block, then grouped NUMERIC metrics (short values only). The long
    #    narrative (description) and the migration-flag prose are deliberately NOT here - they live in
    #    the Blockers sheet and the .txt overview - so this sheet stays clean and scannable.
    # Values are stored as STRINGS so Excel left-aligns them all (numeric values would right-align
    # and float to the far side of the column). Column widths are fixed below so a long URL wraps
    # instead of stretching the value column.
    # THREE columns: Metric (A), Value (B, short numbers - left-aligned), Details (C, long text like
    # the host URL / type). Short values go to B, long narrative to C, so numbers sit right next to
    # their label and the long URL lives in its own wide wrapping column.
    $blank = [pscustomobject]@{ Metric = ''; Value = ''; Details = '' }
    $num = { param($m, $v) [pscustomobject]@{ Metric = $m; Value = [string]$v; Details = '' } }
    $txt = { param($m, $d) [pscustomobject]@{ Metric = $m; Value = ''; Details = $d } }
    $summaryRows = @()
    $summaryRows += (& $txt 'Form' $name)
    if ($ft) { $summaryRows += (& $txt 'Type' $ft.Kind) }
    if ($ft -and $ft.BaseUrl) { $summaryRows += (& $txt 'Location' $ft.BaseUrl) }
    $summaryRows += (& $num 'Complexity' "$score / $bucket")
    if ($formIssues -gt 0) { $summaryRows += (& $txt '** DATA QUALITY **' "$formIssues analyzer issue(s) - some logic may be missing; see _Analysis-Log.txt") }
    $summaryRows += $blank
    $summaryRows += (& $num 'Views' @($views).Count)
    $summaryRows += (& $num 'Sections' $sectionCount)
    $summaryRows += (& $num '   - repeating sections / tables' $repeatingSections)
    $summaryRows += (& $num 'Fields' @($fieldRows).Count)
    $summaryRows += (& $num '   - SharePoint columns' $spBoundCount)
    $summaryRows += (& $num '   - XML-only (not a column)' $xmlOnlyCount)
    $summaryRows += (& $num '   - complex (lookup/choice/person)' $complexCount)
    $summaryRows += (& $num '   - repeating fields' $repeatingFields)
    $summaryRows += (& $num '   - attachment controls' $attachmentCount)
    $summaryRows += (& $num '   - unused (custom)' $unusedCount)
    $summaryRows += $blank
    $summaryRows += (& $num 'Data connections' @($adapters).Count)
    $summaryRows += (& $num '   - secondary' $secondary.Count)
    $summaryRows += (& $num '   - to other SharePoint lists' $otherListIds.Count)
    $summaryRows += (& $num '   - SQL / database' $sqlConns.Count)
    $summaryRows += (& $num '   - web service' $wsConns.Count)
    $summaryRows += $blank
    $summaryRows += (& $num 'Rules' $rulesCount)
    $summaryRows += (& $num '   - with conditions' $condCount)
    $summaryRows += (& $num 'On-load actions' @($onLoad).Count)
    $summaryRows += (& $num 'Field-change handlers' @($events).Count)
    $summaryRows += (& $num 'Calculations' @($calcs).Count)
    $summaryRows += (& $num 'Validation rules' @($valid).Count)
    $summaryRows += (& $num 'Read-only / disable rules' @($readonlyRows).Count)
    $summaryRows += (& $num 'View-navigation rules' $navCount)
    $summaryRows += $blank
    $summaryRows += (& $num 'Code-behind (.dll)' $(if ($files.HasCodeBehind) { 'YES' } else { 'no' }))
    $summaryRows += (& $num 'Browser-enabled' $(if ($browserEnabled) { 'yes' } elseif ($runtime) { 'no (client only)' } else { 'unknown' }))
    $summaryRows += (& $num 'User roles' @($roles.Roles).Count)
    $summaryRows += (& $num 'Migration blockers (see Blockers sheet)' @($blockers).Count)
    $summaryRows += (& $num 'Orphaned files (quarantined in _Unused)' $files.OrphanedTotal)
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Summary' -Rows $summaryRows -NoTable -NoAutoSize

    # 2) Logic (the centrepiece) - WHEN / IF / THEN, grouped by trigger
    $logicSorted = @($logic.Rows | Sort-Object Trigger, RuleSet)
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Logic' -Rows $logicSorted

    # 3) Fields
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Fields' -Rows (@($fieldRows | Sort-Object @{e={$_.Status -ne 'Active'}}, Field))

    # 3b) Structure - the section/field hierarchy (indented), in form order, as a readable blueprint.
    $structRows = @($structure | ForEach-Object {
        $isSec = ($_.Kind -like '*ection*')
        [pscustomobject]@{
            'Section / field' = (('      ' * [Math]::Max(0, $_.Depth - 1)) + $_.Name)
            Kind = $_.Kind
            Label = $(if ($isSec) { '' } else { $labelByField[$_.Name] })
            Control = $(if ($isSec) { '' } else { $controlByField[$_.Name] })
            Type = $_.Type
            Storage = $(if ($isSec) { '' } elseif ($isListForm -or $promoted.ContainsKey($_.Name)) { 'SharePoint' } else { 'XML only' })
            Default = $(if ($isSec) { '' } else { $defaults[$_.Name] })
        }
    })
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Structure' -Rows $structRows -NoTable

    # 4) Connections
    $connRows = @($adapters | ForEach-Object {
        [pscustomobject]@{
            Connection = $_.DataObject; Kind = $_.Kind; Primary = $_.IsPrimary
            OtherList = ($_.IsSharePointList -and $_.ListId -and $_.ListId -ne $primaryListId)
            OnLoad = $_.OnLoad; CanQuery = $_.QueryAllowed; CanSubmit = $_.SubmitAllowed; Detail = Clamp $_.Detail 220
        }
    })
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Connections' -Rows $connRows

    # 5) Views
    $viewRows = @($views | ForEach-Object {
        [pscustomobject]@{
            View = $_.View; File = $_.File; Default = $_.IsDefault
            CanvasAppEquivalent = $(if ($_.IsDefault) { 'Main screen' } else { 'Additional screen / state' })
        }
    })
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Views' -Rows $viewRows

    # 6) Visibility (section show/hide -> Power Fx). Dedupe identical conditions - the same test
    #    (e.g. Ship_To="X") is often repeated across many controls; one row per distinct rule.
    $visibilityRows = @($visibilityRows | Sort-Object View, ShownWhen -Unique)
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Visibility' -Rows $visibilityRows

    # 6b) Read-only / disable conditions (field becomes read-only when ...).
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'ReadOnly' -Rows $readonlyRows

    # 7) Calculations (dynamic defaults)
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Calculations' -Rows $calcs

    # 8) Validation
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Validation' -Rows $valid

    # 9) Dropdowns - grouped by distinct option-set so a form with 80 identical Yes/No/NA
    #    dropdowns collapses to a handful of rows instead of overwhelming the reader.
    $dropGrouped = @()
    foreach ($g in ($dropdownRows | Group-Object Type, Options)) {
        $flds = @($g.Group | ForEach-Object { $_.Field } | Where-Object { $_ } | Select-Object -Unique)
        $fldList = ($flds | Select-Object -First 20) -join ', '
        if ($flds.Count -gt 20) { $fldList += " (+$($flds.Count - 20) more)" }
        $dropGrouped += [pscustomobject]@{
            Type = $g.Group[0].Type; Options = $g.Group[0].Options
            UsedBy = $flds.Count; Fields = $fldList
        }
    }
    $dropGrouped = @($dropGrouped | Sort-Object -Property UsedBy -Descending)
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Dropdowns' -Rows $dropGrouped

    # 10) Navigation (button/event -> view)
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Navigation' -Rows $navigation

    # 11) Roles / identity
    $roleRows = @()
    foreach ($r in $roles.Roles) { $roleRows += [pscustomobject]@{ Item = "Role: $($r.Role)"; Detail = "$($r.Notes) $($r.Members)".Trim() } }
    foreach ($s in $roles.IdentitySignals) { $roleRows += [pscustomobject]@{ Item = 'Identity logic'; Detail = $s } }
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'RolesIdentity' -Rows $roleRows

    # 12) Migration blockers
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Blockers' -Rows $blockers

    # 13) Simplify (dead logic / unused fields)
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Simplify' -Rows $simplify

    # 14) Files
    $pkg = Add-Sheet -Package $pkg -Path $xlsx -Name 'Files' -Rows $files.Rows

    # Make the Summary readable: fixed column widths + wrap, so a long host URL wraps within the
    # value column instead of stretching it and shoving the numbers far to the right. Bold labels.
    try {
        $ws = $pkg.Workbook.Worksheets['Summary']
        if ($ws) {
            $ws.Column(1).Width = 40                                  # Metric label (bold)
            $ws.Column(2).Width = 14                                  # short Value / number
            $ws.Column(3).Width = 72                                  # long Details (URL / type) - wraps
            $ws.Column(1).Style.Font.Bold = $true
            # Force the value column to LEFT-align so numbers sit next to the label instead of
            # right-aligning to the far edge.
            $ws.Column(2).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
            $ws.Column(3).Style.WrapText = $true
            $ws.Cells.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
        }
    } catch {}

    # Finalize: write the build file once, then atomically replace the visible workbook.
    try { Close-ExcelPackage -ExcelPackage $pkg } catch { Log-Issue $name 'workbook' "Finalize failed: $($_.Exception.Message)" 'Error' }
    if (Remove-FileSafe $xlsxFinal) {
        $moved = $false
        for ($mi = 0; $mi -lt 10; $mi++) {
            try { Move-Item -LiteralPath $xlsx -Destination $xlsxFinal -Force; $moved = $true; break }
            catch { Start-Sleep -Milliseconds 800 }
        }
        if (-not $moved) { Log-Issue $name 'workbook' "Built but final file is locked; fresh copy left at $(Split-Path $xlsx -Leaf)" }
    } else {
        Log-Issue $name 'workbook' "'$([System.IO.Path]::GetFileName($xlsxFinal))' is locked; fresh workbook left at $(Split-Path $xlsx -Leaf)"
    }

    # ============================================================================
    #  Write the .txt overview
    # ============================================================================
    $txtPath = Join-Path $folder ("{0}-Overview.txt" -f $name)
    $lines = @()
    $lines += "InfoPath form inventory: $name"
    $lines += ('=' * 60)
    if ($ft) { $lines += "Form type        : $($ft.Kind)"; $lines += "                   $($ft.Description)"; if ($ft.BaseUrl) { $lines += "Host list        : $($ft.BaseUrl)" } }
    $lines += "Complexity       : $score ($bucket)"
    $lines += ""
    $lines += "Views                  : $(@($views).Count)"
    $lines += "Data connections       : $(@($adapters).Count)  (secondary: $($secondary.Count), other SP lists: $($otherListIds.Count), SQL: $($sqlConns.Count), web service: $($wsConns.Count))"
    $lines += "Fields                 : $(@($fieldRows).Count)  (complex: $complexCount, unused: $unusedCount)"
    $lines += "OnLoad triggers        : $(@($onLoad).Count)"
    $lines += "Field-change handlers  : $(@($events).Count)"
    $lines += "Rule sets              : $(@(Sel $root '//xsf:ruleSets/xsf:ruleSet' $nsm).Count)"
    $lines += "Rules                  : $rulesCount  (with conditions: $condCount)"
    $lines += "Calculations           : $(@($calcs).Count)"
    $lines += "Validation rules       : $(@($valid).Count)"
    $lines += "View-navigation rules  : $navCount"
    $lines += "Has code-behind (.dll) : $(if ($files.HasCodeBehind) { 'YES' } else { 'no' })"
    $lines += ""
    if (@($onLoad).Count -gt 0) {
        $lines += "On load:"
        foreach ($o in $onLoad) { $lines += "  - $o" }
        $lines += ""
    }
    $lines += "Files:"
    $lines += "  active   : $(@($files.Rows | Where-Object { $_.Status -eq 'Active' }).Count)"
    $lines += "  orphaned : $(@($files.Rows | Where-Object { $_.Status -eq 'Orphaned' }).Count) (moved to _Unused\)"
    if ($flags.Count -gt 0) {
        $lines += ""
        $lines += "Migration flags:"
        foreach ($fl in $flags) { $lines += "  - $fl" }
    }
    Set-Content -LiteralPath $txtPath -Value $lines -Encoding UTF8

    # ============================================================================
    #  View-flow diagram (Mermaid) + single LLM-context markdown file
    # ============================================================================
    $defaultView = ''
    $dv = $views | Where-Object { $_.IsDefault } | Select-Object -First 1
    if ($dv) { $defaultView = $dv.View }
    $mermaid = Build-ViewFlowMermaid -Views $views -Navigation $navigation -DefaultView $defaultView

    $md = @()
    $md += "# InfoPath form: $name"
    $md += ''
    $md += "> Auto-generated inventory for migration planning. The form is NOT migrated here - this"
    $md += "> file captures its data, logic, dependencies and complexity so a human (or an LLM) can"
    $md += "> decide how to rebuild it. Dead-end files (old schema versions, sample data, images,"
    $md += "> binaries, upgrade transforms) are excluded; the analysis already distils what matters."
    $md += ''
    $md += (To-MdTable -Rows (@($summaryRows | Where-Object { $_.Metric -ne '' } | ForEach-Object {
        [pscustomobject]@{ Metric = $_.Metric; Value = $(if ($_.Value) { $_.Value } else { $_.Details }) }
    })) -Title 'Summary')
    if (@($blockers).Count -gt 0) { $md += (To-MdTable -Rows $blockers -Title 'Migration blockers') }
    $md += '### View flow'
    $md += ''
    $md += $mermaid
    $md += ''
    $md += (To-MdTable -Rows $viewRows -Title 'Views')
    $md += (To-MdTable -Rows $logicSorted -Title 'Logic (when / if / then)')
    if (@($navigation).Count -gt 0) { $md += (To-MdTable -Rows $navigation -Title 'View navigation') }
    $md += (To-MdTable -Rows (@($fieldRows | Sort-Object @{e={$_.Status -ne 'Active'}}, Field)) -Title 'Fields')
    if (@($structure).Count -gt 0) {
        $md += '### Form layout (sections, fields, labels, controls - in form order)'
        $md += ''
        $md += 'Each field line: name  "label"  (control, type)  [SP=SharePoint column / xml=XML-only]  =default'
        $md += ''
        $md += '```'
        foreach ($s in $structure) {
            $line = ('  ' * [Math]::Max(0, $s.Depth)) + $s.Name
            if ($s.Kind -like '*ection*') {
                $line += "  [$($s.Kind)]"
            } else {
                $bits = @()
                $lbl = $labelByField[$s.Name]
                if ($lbl) { $bits += ('"' + $lbl + '"') }
                $tc = @()
                if ($controlByField[$s.Name]) { $tc += $controlByField[$s.Name] }
                if ($s.Type) { $tc += ($s.Type -replace '^xsd:', '') }
                if ($tc.Count -gt 0) { $bits += ('(' + ($tc -join ', ') + ')') }
                if ($isListForm -or $promoted.ContainsKey($s.Name)) { $bits += '[SP]' } else { $bits += '[xml]' }
                if ($defaults.ContainsKey($s.Name)) { $bits += ('=' + $defaults[$s.Name]) }
                if ($bits.Count -gt 0) { $line += '  ' + ($bits -join '  ') }
            }
            $md += $line
        }
        $md += '```'
        $md += ''
    }
    $md += (To-MdTable -Rows $connRows -Title 'Data connections')
    if (@($calcs).Count -gt 0) { $md += (To-MdTable -Rows $calcs -Title 'Calculations / dynamic defaults') }
    if (@($valid).Count -gt 0) { $md += (To-MdTable -Rows $valid -Title 'Validation rules') }
    if (@($visibilityRows).Count -gt 0) { $md += (To-MdTable -Rows $visibilityRows -Title 'Section visibility') }
    if (@($readonlyRows).Count -gt 0) { $md += (To-MdTable -Rows $readonlyRows -Title 'Read-only / disable conditions') }
    if (@($dropGrouped).Count -gt 0) { $md += (To-MdTable -Rows $dropGrouped -Title 'Drop-down / choice options (grouped by option-set)') }
    if (@($roleRows).Count -gt 0) { $md += (To-MdTable -Rows $roleRows -Title 'Roles / identity logic') }
    if (@($simplify).Count -gt 0) { $md += (To-MdTable -Rows $simplify -Title 'Simplification opportunities (dead/unused logic)') }
    $md += (To-MdTable -Rows $files.Rows -Title 'Files (active vs moved to _Unused)')

    # Appendix: a single COMPACT logic-source extract (exact expressions / field paths / GUIDs),
    # stripped of xmlns soup, whitespace, the packaged-file list, schema/extension/design chrome.
    # The distilled tables above already carry the data structure (Fields), so the raw .xsd dumps
    # and the full manifest are deliberately NOT included - this keeps the file small for an LLM.
    $logicXml = ''
    try { $logicXml = Get-CleanLogicXml -Root $root -Nsm $nsm } catch { $logicXml = '' }
    if ($logicXml) {
        if ($logicXml.Length -gt 60000) { $logicXml = $logicXml.Substring(0, 60000) + "`n<!-- truncated -->" }
        $md += '---'
        $md += '## Appendix - Raw logic source (manifest.xsf, namespaces and boilerplate stripped)'
        $md += ''
        $md += 'Only the logic-bearing elements (views, submit, data connections, rule sets, event'
        $md += 'handlers, on-load, calculations, validation). Exact expressions kept for fidelity.'
        $md += ''
        $md += '```xml'
        $md += $logicXml
        $md += '```'
    }

    $mdPath = Join-Path $folder 'LLM Context.md'
    Set-Content -LiteralPath $mdPath -Value ($md -join "`n") -Encoding UTF8

    # All cleaned view layouts (xmlns/CSS/geometry stripped) combined into ONE companion file, kept
    # separate from the curated LLM Context.md so it does not bloat it.
    $vmd = @()
    $vmd += "# $name - cleaned view layouts (additional LLM context)"
    $vmd += ''
    $vmd += '> Each view''s .xsl with namespaces, CSS and geometry attributes stripped, but table'
    $vmd += '> structure, control bindings (xd:binding), labels and xsl:if/choose conditions kept.'
    $vmd += '> Use alongside "LLM Context.md" when you need the exact on-screen layout.'
    $vmd += ''
    $wroteAny = $false
    foreach ($v in $views) {
        if (-not $v.File) { continue }
        $vp = Join-Path $folder $v.File
        if (-not (Test-Path -LiteralPath $vp)) { continue }
        $vt = ''
        try { $vt = [System.IO.File]::ReadAllText($vp) } catch { continue }
        $vmd += "## View: $($v.View)  ($($v.File))"
        $vmd += ''
        $vmd += '```xml'
        $vmd += (Clean-XslText $vt)
        $vmd += '```'
        $vmd += ''
        $wroteAny = $true
    }
    if ($wroteAny) {
        $vmdPath = Join-Path $folder 'LLM Context - additional XSL files.md'
        try { Set-Content -LiteralPath $vmdPath -Value ($vmd -join "`n") -Encoding UTF8 } catch {}
    }

    Write-Host ("    -> {0}  (complexity {1}, {2}; moved {3} junk file(s))" -f (Split-Path $xlsxFinal -Leaf), $score, $bucket, $files.MovedCount)

    # Roll-up row
    return [pscustomobject]@{
        Form = $name
        Type = $(if ($ft) { $ft.Kind } else { 'Unknown' })
        HostList = $(if ($ft) { $ft.BaseUrl } else { '' })
        Complexity = $score
        Bucket = $bucket
        Views = @($views).Count
        Fields = @($fieldRows).Count
        SharePointColumns = $spBoundCount
        XmlOnlyFields = $xmlOnlyCount
        UnusedFields = $unusedCount
        ComplexFields = $complexCount
        Sections = $sectionCount
        RepeatingSections = $repeatingSections
        RepeatingFields = $repeatingFields
        Attachments = $attachmentCount
        Connections = @($adapters).Count
        OtherSPLists = $otherListIds.Count
        SqlConns = $sqlConns.Count
        WebServiceConns = $wsConns.Count
        Rules = $rulesCount
        Conditions = $condCount
        OnLoadActions = @($onLoad).Count
        FieldChangeHandlers = @($events).Count
        Calculations = @($calcs).Count
        Validations = @($valid).Count
        ReadOnlyRules = @($readonlyRows).Count
        Navigation = $navCount
        CodeBehind = $files.HasCodeBehind
        UserRoles = @($roles.Roles).Count
        Blockers = @($blockers).Count
        BrowserEnabled = $browserEnabled
        OrphanedFiles = $files.OrphanedTotal
    }
}

# ============================================================================================
#  Main
# ============================================================================================

Write-Host ""
Write-Host "InfoPath XSN form analyzer"
Write-Host "=========================="
Write-Host "Folder: $scriptRoot"
Write-Host ""

# ImportExcel - auto-install for the current user if it is not already present.
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found - installing it for the current user (one-time)..."
    try {
        # SharePoint Gallery needs TLS 1.2 on older Windows configs.
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
        }
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "ImportExcel installed."
    } catch {
        Write-Warning "Could not auto-install ImportExcel: $($_.Exception.Message)"
        Write-Warning "Install it manually once, then re-run:  Install-Module ImportExcel -Scope CurrentUser"
        exit 1
    }
}
try {
    Import-Module ImportExcel -ErrorAction Stop
} catch {
    Write-Warning "ImportExcel is present but could not be loaded: $($_.Exception.Message)"
    exit 1
}

# Find xsn files
$xsnFiles = @(Get-ChildItem -LiteralPath $scriptRoot -Filter '*.xsn' -File -ErrorAction SilentlyContinue | Sort-Object Name)
if ($xsnFiles.Count -eq 0) {
    Write-Warning "No .xsn files found in $scriptRoot. Put this script in the folder with your .xsn files."
    exit 1
}

# Process every .xsn found in the folder, automatically (no prompt).
$selected = $xsnFiles
Write-Host "Found $($xsnFiles.Count) .xsn file(s) - analyzing all of them:"
for ($i = 0; $i -lt $xsnFiles.Count; $i++) {
    Write-Host ("  [{0,2}] {1}" -f ($i + 1), $xsnFiles[$i].Name)
}
Write-Host ""

$rollup = @()
for ($i = 0; $i -lt $selected.Count; $i++) {
    $f = $selected[$i]
    $pct = [int](100.0 * ($i + 1) / $selected.Count)
    Write-Progress -Activity 'Analyzing XSN forms' -Status ("[{0}/{1}] {2}" -f ($i + 1), $selected.Count, $f.Name) -PercentComplete $pct
    $row = $null
    try {
        $row = Analyze-Form -XsnPath $f.FullName
    } catch {
        Log-Issue ([System.IO.Path]::GetFileNameWithoutExtension($f.Name)) 'analyze' "$($_.Exception.GetType().FullName): $($_.Exception.Message)" 'Error'
    }
    if ($row) { $rollup += $row }
}
Write-Progress -Activity 'Analyzing XSN forms' -Completed

# Which selected forms produced NO roll-up row (i.e. were skipped / failed entirely)?
$analyzedNames = @($rollup | ForEach-Object { $_.Form })
foreach ($f in $selected) {
    $bn = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    if ($analyzedNames -notcontains $bn) {
        if (-not (@($script:RunLog | Where-Object { $_.Form -eq $bn -and $_.Severity -eq 'Error' }).Count)) {
            Log-Issue $bn 'analyze' 'Form produced no report row for an unknown reason.' 'Error'
        }
    }
}

# Roll-up workbook
if ($rollup.Count -gt 0) {
    $rollupPath = Join-Path $scriptRoot '_AllForms-Summary.xlsx'
    $rollupTmp = Join-Path $scriptRoot '~_AllForms-Summary.build.xlsx'
    [void](Remove-FileSafe $rollupTmp)
    try {
        $rollup | Sort-Object -Property Complexity -Descending |
            Export-Excel -Path $rollupTmp -WorksheetName 'AllForms' -TableName 'AllForms' -AutoSize -FreezeTopRow -TableStyle 'Medium2'
        if (Remove-FileSafe $rollupPath) {
            $moved = $false
            for ($mi = 0; $mi -lt 10; $mi++) {
                try { Move-Item -LiteralPath $rollupTmp -Destination $rollupPath -Force; $moved = $true; break }
                catch { Start-Sleep -Milliseconds 800 }
            }
            if ($moved) { Write-Host ""; Write-Host "Roll-up written: $rollupPath" }
            else { Log-Issue '(roll-up)' 'rollup' "Built but _AllForms-Summary.xlsx is locked; fresh copy left at $(Split-Path $rollupTmp -Leaf)" }
        } else {
            Log-Issue '(roll-up)' 'rollup' "_AllForms-Summary.xlsx is locked (OneDrive/Excel); fresh copy left at $(Split-Path $rollupTmp -Leaf)"
        }
    } catch {
        Log-Issue '(roll-up)' 'rollup' "Could not write the roll-up workbook: $($_.Exception.Message)" 'Error'
    }
}

# ----------------------------------------------------------------------------------------------
#  Write the run log so any failure is visible and no missing logic goes unnoticed.
# ----------------------------------------------------------------------------------------------
$logPath = Join-Path $scriptRoot '_Analysis-Log.txt'
$errors = @($script:RunLog | Where-Object { $_.Severity -eq 'Error' })
$warns = @($script:RunLog | Where-Object { $_.Severity -ne 'Error' })
$skipped = @($selected | Where-Object { $analyzedNames -notcontains [System.IO.Path]::GetFileNameWithoutExtension($_.Name) })
$logLines = @()
$logLines += "InfoPath XSN analysis - run log"
$logLines += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$logLines += "Forms selected: $($selected.Count)   fully analyzed: $($rollup.Count)   skipped/failed: $($skipped.Count)"
$logLines += "Errors: $($errors.Count)   Warnings: $($warns.Count)"
$logLines += ('=' * 70)
if ($script:RunLog.Count -eq 0) {
    $logLines += ""
    $logLines += "No issues. Every selected form parsed cleanly and every analyzer completed."
    $logLines += "Nothing is missing from the reports."
} else {
    if ($errors.Count -gt 0) {
        $logLines += ""
        $logLines += "ERRORS (these forms / stages did NOT produce complete output):"
        foreach ($e in $errors) { $logLines += "  [ERROR] $($e.Form) :: $($e.Stage) :: $($e.Message)" }
    }
    if ($warns.Count -gt 0) {
        $logLines += ""
        $logLines += "WARNINGS (output produced, but review):"
        foreach ($w in $warns) { $logLines += "  [warn]  $($w.Form) :: $($w.Stage) :: $($w.Message)" }
    }
}
try { Set-Content -LiteralPath $logPath -Value $logLines -Encoding UTF8 } catch {}

Write-Host ""
Write-Host "Done. Analyzed $($rollup.Count) of $($selected.Count) form(s)."
if ($errors.Count -gt 0) {
    Write-Host ("ISSUES: {0} error(s), {1} warning(s). See {2}" -f $errors.Count, $warns.Count, (Split-Path $logPath -Leaf)) -ForegroundColor Yellow
    exit 1
} else {
    Write-Host ("No errors. {0} warning(s). Log: {1}" -f $warns.Count, (Split-Path $logPath -Leaf))
    exit 0
}
