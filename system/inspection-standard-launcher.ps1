$ErrorActionPreference = 'Stop'

# === Schema versions ===
$CacheSchemaVersion = '2.0'
$ParserVersion = '1.0'

# === CLI arguments ===
$forceRebuild   = $args -contains '--force-rebuild'
$clearCache     = $args -contains '--clear-cache'
$prepareOnly    = $args -contains '--prepare-only'

# === Root detection ===
$ScriptDirectory = $PSScriptRoot
if (-not $ScriptDirectory) {
  $ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$AppRoot = $null
$rootCandidate = Get-Item -LiteralPath $ScriptDirectory
while ($null -ne $rootCandidate) {
  $candidateHtml = Join-Path $rootCandidate.FullName 'Inspection Standard Tool.html'
  $candidateTemplate = Join-Path $rootCandidate.FullName 'Inspection Standard Template.xlsx'
  if ((Test-Path -LiteralPath $candidateHtml -PathType Leaf) -and
      (Test-Path -LiteralPath $candidateTemplate -PathType Leaf)) {
    $AppRoot = $rootCandidate.FullName
    break
  }
  $rootCandidate = $rootCandidate.Parent
}

if (-not $AppRoot) {
  throw "Project root not found. Expected both 'Inspection Standard Tool.html' and 'Inspection Standard Template.xlsx' above '$ScriptDirectory'."
}

$ToolsParent = Split-Path $AppRoot -Parent
$systemDir = Join-Path $AppRoot 'system'

# === Directory setup ===
$generatedDir = Join-Path $systemDir 'generated'
$cacheDir     = Join-Path $systemDir 'cache'
$logDir       = Join-Path $systemDir 'logs'
$chromeProfileDir = Join-Path $cacheDir 'browser-profile'

foreach ($d in @($generatedDir, $cacheDir, $logDir, $chromeProfileDir)) {
  if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

# === Logging ===
$logPath = Join-Path $logDir 'Inspection_Standard_Tool_launcher.log'

# Rotate log if > 1 MB
if ((Test-Path -LiteralPath $logPath) -and ((Get-Item -LiteralPath $logPath).Length -gt 1MB)) {
  $archivePath = Join-Path $logDir ('Inspection_Standard_Tool_launcher.{0:yyyyMMdd-HHmmss}.log' -f (Get-Date))
  Move-Item -LiteralPath $logPath -Destination $archivePath -Force
}

$script:logStart = Get-Date
function Write-Log {
  param([string]$msg)
  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $line = "$ts $msg"
  Write-Host $line
  Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}

Write-Log "=== Inspection Standard Tool Launcher v$CacheSchemaVersion (parser v$ParserVersion) ==="
Write-Log "Script Directory: $ScriptDirectory"
Write-Log "Application Root: $AppRoot"
Write-Log "Tools Parent: $ToolsParent"
Write-Log "Arguments: forceRebuild=$forceRebuild, clearCache=$clearCache, prepareOnly=$prepareOnly"

# === Embedded C# XLSX reader (shared from Master Carton launcher) ===
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
if (-not ('WvnXlsxReader' -as [type])) {
  Add-Type -ReferencedAssemblies 'System.IO.Compression.dll','System.IO.Compression.FileSystem.dll','System.Xml.dll' -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml;
public sealed class WvnSheetRows { public string Name; public List<string[]> Rows = new List<string[]>(); }
public static class WvnXlsxReader {
  static string EntryText(ZipArchive z, string name) { var e=z.GetEntry(name); if(e==null)return null; using(var s=e.Open())using(var r=new StreamReader(s))return r.ReadToEnd(); }
  static int Col(string reference) { int n=0; foreach(char c in reference){ if(c<'A'||c>'Z')break; n=n*26+(c-'A'+1); } return n-1; }
  public static List<WvnSheetRows> Read(string path, string pattern, int skip) {
    var output=new List<WvnSheetRows>();
    using(var fs=File.OpenRead(path))using(var z=new ZipArchive(fs,ZipArchiveMode.Read,false)) {
      var shared=new List<string>(); string st=EntryText(z,"xl/sharedStrings.xml");
      if(st!=null){var d=new XmlDocument();d.LoadXml(st);foreach(XmlNode n in d.SelectNodes("//*[local-name()='si']"))shared.Add(n.InnerText);}
      var wb=new XmlDocument();wb.LoadXml(EntryText(z,"xl/workbook.xml")); var rd=new XmlDocument();rd.LoadXml(EntryText(z,"xl/_rels/workbook.xml.rels"));
      var rel=new Dictionary<string,string>();foreach(XmlNode n in rd.SelectNodes("//*[local-name()='Relationship']"))rel[n.Attributes["Id"].Value]=n.Attributes["Target"].Value;
      int number=0; foreach(XmlNode sn in wb.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']")) {
        if(number++<skip)continue; string name=sn.Attributes["name"].Value;if(!Regex.IsMatch(name,pattern))continue;
        string rid=sn.Attributes["r:id"].Value,target;if(!rel.TryGetValue(rid,out target))continue; target=target.TrimStart('/');if(!target.StartsWith("xl/"))target="xl/"+target.TrimStart('.','/');
        var entry=z.GetEntry(target);if(entry==null)continue;var sheet=new WvnSheetRows();sheet.Name=name;
        using(var stream=entry.Open())using(var xr=XmlReader.Create(stream,new XmlReaderSettings{IgnoreComments=true,IgnoreWhitespace=true})) {
          while(xr.Read())if(xr.NodeType==XmlNodeType.Element&&xr.LocalName=="row") {
            var cells=new Dictionary<int,string>();int max=-1;
            using(var rr=xr.ReadSubtree())while(rr.Read())if(rr.NodeType==XmlNodeType.Element&&rr.LocalName=="c") {
              string cr=rr.GetAttribute("r")??"A",type=rr.GetAttribute("t")??"",value="";int ci=Col(cr.ToUpperInvariant());
              using(var cc=rr.ReadSubtree())while(cc.Read())if(cc.NodeType==XmlNodeType.Element&&(cc.LocalName=="v"||cc.LocalName=="t"))value=cc.ReadElementContentAsString();
              int si;if(type=="s"&&Int32.TryParse(value,out si)&&si>=0&&si<shared.Count)value=shared[si];cells[ci]=(value??"").Trim();if(ci>max)max=ci;
            }
            var row=new string[Math.Max(0,max+1)];foreach(var kv in cells)row[kv.Key]=kv.Value;for(int i=0;i<row.Length;i++)if(row[i]==null)row[i]="";sheet.Rows.Add(row);
          }
        }
        output.Add(sheet);
      }
    }
    return output;
  }
}
'@
}

# === Shared helper functions ===
function Normalize-Header { param($value) return ([string]$value).Trim().ToLowerInvariant() -replace '[^a-z0-9]+', ' ' -replace '^\s+|\s+$', '' }
function Get-RowCell { param($row, [int]$index) if ($index -lt 0 -or $null -eq $row -or $index -ge $row.Count) { return '' }; return ([string]$row[$index]).Trim() }
function Normalize-ItemId { param($Value) if ($null -eq $Value) { return '' }; $text = ([string]$Value).Trim(); if ($text -match '^\d+\.0+$') { return $text -replace '\.0+$', '' }; return $text }

function Read-IniFile {
  param([string]$FilePath)
  $ini = @{}
  if (-not (Test-Path -LiteralPath $FilePath)) { return $ini }
  $section = ""
  foreach ($line in Get-Content -LiteralPath $FilePath) {
    $line = $line.Trim()
    if ($line.StartsWith(';') -or $line.StartsWith('#') -or -not $line) { continue }
    if ($line.StartsWith('[') -and $line.EndsWith(']')) {
      $section = $line.Substring(1, $line.Length - 2).Trim()
      $ini[$section] = @{}
    } elseif ($line -match '^([^=]+)=(.*)$') {
      $key = $matches[1].Trim()
      $val = $matches[2].Trim()
      if ($section) { $ini[$section][$key] = $val }
    }
  }
  return $ini
}

function Get-HeaderInfo {
  param($Sheet, [object[]]$RequiredAliasGroups)
  for ($rowIndex = 0; $rowIndex -lt [math]::Min(40, $Sheet.Rows.Count); $rowIndex++) {
    $map = @{}
    for ($columnIndex = 0; $columnIndex -lt $Sheet.Rows[$rowIndex].Count; $columnIndex++) {
      $key = Normalize-Header $Sheet.Rows[$rowIndex][$columnIndex]
      if ($key -and -not $map.ContainsKey($key)) { $map[$key] = $columnIndex }
    }
    $complete = $true
    foreach ($aliases in $RequiredAliasGroups) {
      $found = $false
      foreach ($alias in @($aliases)) { if ($map.ContainsKey((Normalize-Header $alias))) { $found = $true; break } }
      if (-not $found) { $complete = $false; break }
    }
    if ($complete) { return [pscustomobject]@{ Row = $rowIndex; Map = $map } }
  }
  return $null
}

function Find-HeaderColumn {
  param($HeaderInfo, [string[]]$Aliases)
  if (-not $HeaderInfo) { return -1 }
  foreach ($alias in $Aliases) { $key = Normalize-Header $alias; if ($HeaderInfo.Map.ContainsKey($key)) { return [int]$HeaderInfo.Map[$key] } }
  return -1
}

function Test-IsValidBomWorkbook {
  param([System.IO.FileInfo]$File)
  try {
    $sheets = @([WvnXlsxReader]::Read($File.FullName, '.*', 0))
    $hasBomSheet = $false
    foreach ($sheet in $sheets) {
      for ($hr = 0; $hr -lt [math]::Min(60, $sheet.Rows.Count); $hr++) {
        $hh = @($sheet.Rows[$hr] | ForEach-Object { Normalize-Header $_ })
        $fgIdx = -1; $itemIdx = -1; $productIdx = -1
        for ($hc = 0; $hc -lt $hh.Count; $hc++) {
          if ($hh[$hc] -in @('fg','fg id','fgid','finished good','finishedgood','top item id')) { $fgIdx=$hc }
          if ($hh[$hc] -in @('item number','item no','itemnumber','component','component item','child item id')) { $itemIdx=$hc }
          if ($hh[$hc] -in @('product name','product','productname','description','item description','child item name')) { $productIdx=$hc }
        }
        if ($fgIdx -ge 0 -and $itemIdx -ge 0 -and $productIdx -ge 0) { $hasBomSheet = $true; break }
      }
      if ($hasBomSheet) { break }
    }
    return $hasBomSheet
  } catch {
    return $false
  }
}

function Get-BomBusinessDateFromFilename {
  param([Parameter(Mandatory=$true)][System.IO.FileInfo]$File)
  $res = [pscustomobject]@{ Success=$false; BusinessDate=[DateTime]::MinValue; MonthToken=""; Day=0; InferredYear=0; Error="" }
  $MonthNumbers = @{ JAN=1;FEB=2;MAR=3;APR=4;MAY=5;JUN=6;JUL=7;AUG=8;SEP=9;OCT=10;NOV=11;DEC=12 }
  $name = $File.Name
  if ($name -match '^BOM\s*&\s*FORMULA\s+([A-Za-z]{3})\s+(\d{1,2})\s+D365\.(xlsx|xlsm)$') {
    $monthToken = $Matches[1].ToUpperInvariant()
    $dayVal = [int]$Matches[2]
    if (-not $MonthNumbers.ContainsKey($monthToken)) { $res.Error = "Unrecognized month token '$monthToken'"; return $res }
    $monthVal = $MonthNumbers[$monthToken]
    $mYear = $File.LastWriteTimeUtc.Year
    $candidateDate = $null
    try { $candidateDate = [DateTime]::new([int]$mYear, [int]$monthVal, [int]$dayVal, 0, 0, 0, [System.DateTimeKind]::Utc) }
    catch { try { $candidateDate = [DateTime]::new(([int]$mYear - 1), [int]$monthVal, [int]$dayVal, 0, 0, 0, [System.DateTimeKind]::Utc); $mYear = $mYear - 1 } catch { $res.Error="Invalid date"; return $res } }
    $diffDays = ($candidateDate - $File.LastWriteTimeUtc.Date).TotalDays
    if ($diffDays -gt 45) { $mYear = $mYear - 1; try { $candidateDate = [DateTime]::new([int]$mYear, [int]$monthVal, [int]$dayVal, 0, 0, 0, [System.DateTimeKind]::Utc) } catch { $res.Error="Invalid date after year adjustment"; return $res } }
    $res.Success=$true; $res.BusinessDate=$candidateDate; $res.MonthToken=$monthToken; $res.Day=$dayVal; $res.InferredYear=$mYear
  } else { $res.Error = "Filename does not match BOM & FORMULA pattern" }
  return $res
}

# === Safe JSON write with atomic tmp replacement ===
function Write-AtomicJson {
  param([string]$FilePath, [string]$JsonContent)
  $tmpPath = "$FilePath.tmp"
  try {
    # Write to temp file (UTF-8 without BOM for clean JSON)
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($tmpPath, $JsonContent, $utf8NoBom)
    # Verify temp file exists and has content
    if (-not (Test-Path -LiteralPath $tmpPath)) { throw "Temp file was not created: $tmpPath" }
    if ((Get-Item -LiteralPath $tmpPath).Length -eq 0 -and $JsonContent.Length -gt 0) { throw "Temp file is empty: $tmpPath" }
    # Verify valid JSON by parsing
    try { $null = Get-Content -Raw -LiteralPath $tmpPath | ConvertFrom-Json } catch { throw "Temp file is not valid JSON: $_" }
    # Atomically replace
    Move-Item -LiteralPath $tmpPath -Destination $FilePath -Force
  } catch {
    if (Test-Path -LiteralPath $tmpPath) { Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue }
    throw
  }
}

function Write-AtomicIni {
  param([string]$FilePath, [string[]]$Lines)
  $tmpPath = "$FilePath.tmp"
  try {
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($tmpPath, $Lines, $utf8NoBom)
    if (-not (Test-Path -LiteralPath $tmpPath)) { throw "Temp file was not created: $tmpPath" }
    Move-Item -LiteralPath $tmpPath -Destination $FilePath -Force
  } catch {
    if (Test-Path -LiteralPath $tmpPath) { Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue }
    throw
  }
}

# === Cache file paths ===
$cacheStatePath      = Join-Path $cacheDir 'inspection_bom_cache_state.ini'
$manifestPath        = Join-Path $generatedDir 'inspection_bom_manifest.json'
$preparedDataPath    = Join-Path $generatedDir 'inspection_bom_data.json'
$stagedWorkbookPath  = Join-Path $generatedDir 'inspection_bom_source.xlsx'
$pictureIndexPath    = Join-Path $generatedDir 'master_carton_picture_index.json'
$pictureCacheDir     = Join-Path $cacheDir 'master-carton-pictures'
# Legacy manifest at root for backward compatibility (the HTML also checks here)
$legacyManifestPath  = Join-Path $AppRoot 'Weekly BOM Manifest.json'

if (-not (Test-Path -LiteralPath $pictureCacheDir)) {
  New-Item -ItemType Directory -Path $pictureCacheDir -Force | Out-Null
}

# === --clear-cache ===
if ($clearCache) {
  Write-Log "Clear-cache requested. Removing Inspection Standard cache files..."
  @($cacheStatePath, $manifestPath, $preparedDataPath, $stagedWorkbookPath, $legacyManifestPath) | ForEach-Object {
    if (Test-Path -LiteralPath $_) {
      Remove-Item -LiteralPath $_ -Force -ErrorAction SilentlyContinue
      Write-Log "  Removed: $_"
    }
  }
  Write-Log "Cache cleared."
  if (-not $forceRebuild -and $prepareOnly) {
    Write-Log "Done (--clear-cache --prepare-only, no rebuild)."
    exit 0
  }
}

# === Configuration ===
$configPath = Join-Path $systemDir 'config.ini'
$config = Read-IniFile -FilePath $configPath
Write-Log "Configuration file: $configPath"

if ($config.ContainsKey('Launcher') -and
    $config['Launcher'].ContainsKey('ForceRebuild') -and
    $config['Launcher']['ForceRebuild'] -eq 'true') {
  $forceRebuild = $true
  Write-Log "ForceRebuild enabled by config.ini."
}

$bomPattern = 'BOM & FORMULA*D365.xlsx'
if ($config.ContainsKey('Files') -and $config['Files'].ContainsKey('BomPattern') -and $config['Files']['BomPattern']) {
  $bomPattern = $config['Files']['BomPattern']
}

# === BOM directory discovery (same priority as Master Carton launcher) ===
$configuredWeeklyBomFolder = $null
if ($config.ContainsKey('Search')) {
  if ($config['Search'].ContainsKey('WeeklyBomFolder') -and $config['Search']['WeeklyBomFolder']) {
    $val = $config['Search']['WeeklyBomFolder']
    $resolvedPath = $val
    if (-not [System.IO.Path]::IsPathRooted($val)) {
      $resolvedPath = [System.IO.Path]::GetFullPath((Join-Path $AppRoot $val))
    }
    if (Test-Path -LiteralPath $resolvedPath) {
      $configuredWeeklyBomFolder = $resolvedPath
      Write-Log "Configured WeeklyBomFolder: $configuredWeeklyBomFolder"
    } else {
      Write-Log "WARNING: Configured WeeklyBomFolder '$val' (resolved: '$resolvedPath') does not exist."
    }
  }
  if (-not $configuredWeeklyBomFolder) {
    for ($i = 1; $i -le 10; $i++) {
      $key = "Root$i"
      if ($config['Search'].ContainsKey($key) -and $config['Search'][$key]) {
        $val = $config['Search'][$key]
        $resolvedPath = $val
        if (-not [System.IO.Path]::IsPathRooted($val)) {
          $resolvedPath = [System.IO.Path]::GetFullPath((Join-Path $AppRoot $val))
        }
        if (Test-Path -LiteralPath $resolvedPath) {
          $configuredWeeklyBomFolder = $resolvedPath
          Write-Log "Configured Root$i override: $configuredWeeklyBomFolder"
          break
        }
      }
    }
  }
}

$resolvedWeeklyBomDir = $null
$siblingWeeklyBomDir = Join-Path $ToolsParent '1. Weekly BOM'
$siblingDiscoverySucceeded = $false

if ($configuredWeeklyBomFolder) {
  $resolvedWeeklyBomDir = $configuredWeeklyBomFolder
  Write-Log "Using configured path override: $resolvedWeeklyBomDir"
} else {
  # Priority 1: Sibling directory (parent of tool / "1. Weekly BOM")
  if (Test-Path -LiteralPath $siblingWeeklyBomDir) {
    $resolvedWeeklyBomDir = (Get-Item -LiteralPath $siblingWeeklyBomDir).FullName
    $siblingDiscoverySucceeded = $true
    Write-Log "Sibling Weekly BOM directory found: $resolvedWeeklyBomDir"
  } else {
    Write-Log "Sibling directory not found: $siblingWeeklyBomDir"
    # Priority 2: Search parent for '1. Weekly BOM' (non-recursive)
    Write-Log "Searching '$ToolsParent' for '1. Weekly BOM'..."
    $foundDirs = Get-ChildItem -LiteralPath $ToolsParent -Directory -Filter '1. Weekly BOM' -ErrorAction SilentlyContinue
    if ($foundDirs.Count -gt 0) {
      $resolvedWeeklyBomDir = $foundDirs[0].FullName
      $siblingDiscoverySucceeded = $true
      Write-Log "Discovered '1. Weekly BOM' under parent: $resolvedWeeklyBomDir"
    } else {
      # Priority 3: Check adjacent Master Carton tool directories
      $mcDirs = Get-ChildItem -LiteralPath $ToolsParent -Directory -Filter 'Mastercarton check tool*' -ErrorAction SilentlyContinue
      $mcBomFound = $false
      foreach ($mcDir in $mcDirs) {
        $mcBomDir = Join-Path $mcDir.FullName '1. Weekly BOM'
        if (Test-Path -LiteralPath $mcBomDir) {
          $resolvedWeeklyBomDir = (Get-Item -LiteralPath $mcBomDir).FullName
          Write-Log "Found '1. Weekly BOM' inside Master Carton tool: $resolvedWeeklyBomDir"
          $mcBomFound = $true
          break
        }
      }
      if (-not $mcBomFound) {
        # Priority 4: Local fallback
        $localBomDir = Join-Path $AppRoot '1. Weekly BOM'
        if (Test-Path -LiteralPath $localBomDir) {
          $resolvedWeeklyBomDir = (Get-Item -LiteralPath $localBomDir).FullName
          Write-Log "Using local fallback: $resolvedWeeklyBomDir"
        }
      }
    }
  }
}

Write-Log "Search summary:"
Write-Log "  Sibling path: $siblingWeeklyBomDir (found: $siblingDiscoverySucceeded)"
Write-Log "  Configured override: $(if($configuredWeeklyBomFolder){$configuredWeeklyBomFolder}else{'[None]'})"
Write-Log "  Resolved BOM directory: $(if($resolvedWeeklyBomDir){$resolvedWeeklyBomDir}else{'[None]'})"

$masterCartonPictureFolderName = 'Master_Carton_Pics'
if ($config.ContainsKey('Search') -and
    $config['Search'].ContainsKey('MasterCartonPictureFolder') -and
    $config['Search']['MasterCartonPictureFolder']) {
  $masterCartonPictureFolderName = $config['Search']['MasterCartonPictureFolder']
}

$resolvedMasterCartonPictureDir = $null
if ($resolvedWeeklyBomDir) {
  if ([System.IO.Path]::IsPathRooted($masterCartonPictureFolderName)) {
    $resolvedMasterCartonPictureDir = $masterCartonPictureFolderName
  } else {
    $resolvedMasterCartonPictureDir = [System.IO.Path]::GetFullPath(
      (Join-Path $resolvedWeeklyBomDir $masterCartonPictureFolderName)
    )
  }
}
Write-Log "  Master-carton pictures: $(if($resolvedMasterCartonPictureDir){$resolvedMasterCartonPictureDir}else{'[None]'})"

# === BOM workbook enumeration and selection ===
Write-Log "Phase: Enumerating Weekly BOM files"

$bomCandidates = @()
$decoratedCandidates = [System.Collections.ArrayList]::new()
$selectedBom = $null
$selectedBomName = ""
$selectedBomModifiedUtc = ""
$selectedBomSize = 0
$bomFilesFound = 0

if ($resolvedWeeklyBomDir -and (Test-Path -LiteralPath $resolvedWeeklyBomDir)) {
  $rawCandidates = Get-ChildItem -LiteralPath $resolvedWeeklyBomDir -File -ErrorAction SilentlyContinue |
    Where-Object {
      $_.Name -notlike '~$*' -and
      $_.Length -gt 0 -and
      ($_.Extension -eq '.xlsx' -or $_.Extension -eq '.xlsm') -and
      $_.Name -like $bomPattern
    }

  $bomFilesFound = @($rawCandidates).Count
  Write-Log "Found $bomFilesFound raw BOM candidate(s) matching pattern '$bomPattern'"

  # Decorate candidates
  foreach ($candidate in $rawCandidates) {
    $parsed = Get-BomBusinessDateFromFilename -File $candidate
    if (-not $parsed.Success) {
      Write-Log "  Filename date could not be parsed for '$($candidate.Name)': $($parsed.Error)"
    }
    [void]$decoratedCandidates.Add([pscustomobject]@{
      File               = $candidate
      Name               = $candidate.Name
      ParsedBusinessDate = $parsed.BusinessDate
      FilenameDateValid  = $parsed.Success
      LastWriteTimeUtc   = $candidate.LastWriteTimeUtc
      Length             = $candidate.Length
      WorkbookValid      = $null
      ValidationError    = ""
    })
  }

  # Sort: filename date valid desc, parsed date desc, LastWriteTimeUtc desc, name desc
  $sortedCandidates = @($decoratedCandidates | Sort-Object `
    @{ Expression='FilenameDateValid'; Descending=$true },
    @{ Expression='ParsedBusinessDate'; Descending=$true },
    @{ Expression='LastWriteTimeUtc'; Descending=$true },
    @{ Expression='Name'; Descending=$true })

  # Validate and select
  foreach ($candidate in $sortedCandidates) {
    Write-Log "  Candidate: $($candidate.Name) ($($candidate.Length) bytes, modified $($candidate.LastWriteTimeUtc.ToString('o')))"
    try {
      $stream = [System.IO.File]::OpenRead($candidate.File.FullName)
      $stream.Close()
      if (Test-IsValidBomWorkbook -File $candidate.File) {
        $candidate.WorkbookValid = $true
      } else {
        $candidate.WorkbookValid = $false
        $candidate.ValidationError = "Failed structural BOM validation."
        Write-Log "    Rejected: $($candidate.ValidationError)"
      }
    } catch {
      $candidate.WorkbookValid = $false
      $candidate.ValidationError = "File read/availability failure: $_"
      Write-Log "    Rejected: $($candidate.ValidationError)"
    }
  }

  $selectedReason = ""
  foreach ($candidate in $sortedCandidates) {
    if ($candidate.WorkbookValid) {
      $selectedBom = $candidate.File
      $selectedBomName = $candidate.Name
      $selectedBomModifiedUtc = $candidate.LastWriteTimeUtc.ToString('o')
      $selectedBomSize = $candidate.Length
      if ($candidate.FilenameDateValid) {
        $selectedReason = "Latest valid business date from filename"
      } else {
        $selectedReason = "Modification-time fallback (filename date not parseable)"
      }
      Write-Log "Selected BOM: $selectedBomName ($selectedReason)"
      break
    }
  }
} else {
  Write-Log "No Weekly BOM directory found."
}

# === Cache freshness check ===
Write-Log "Phase: Cache freshness check"

$cacheStatus = 'unavailable'
$rebuildNeeded = $true
$rebuildReason = ''
$sourceHash = ''
$manifestData = $null
$preparedData = $null

if ($selectedBom) {
  # Fast metadata check
  if (-not $forceRebuild -and (Test-Path -LiteralPath $cacheStatePath)) {
    $cacheState = Read-IniFile -FilePath $cacheStatePath
    if ($cacheState.ContainsKey('BomWorkbook')) {
      $c = $cacheState['BomWorkbook']
      $cPath = $c['SourceFullPath']
      $cLength = $c['SourceLength']
      $cModified = $c['LastWriteTimeUtc']
      $cSchema = $c['CacheSchemaVersion']
      $cParser = $c['ParserVersion']
      $cHash = $c['Hash']

      $fastMatch = ($cPath -eq $selectedBom.FullName) -and
                   ($cLength -eq $selectedBomSize.ToString()) -and
                   ($cModified -eq $selectedBomModifiedUtc) -and
                   ($cSchema -eq $CacheSchemaVersion) -and
                   ($cParser -eq $ParserVersion)

      Write-Log "Fast check: path=$(if($cPath -eq $selectedBom.FullName){'match'}else{'MISMATCH'}), size=$(if($cLength -eq $selectedBomSize.ToString()){'match'}else{'MISMATCH'}), modified=$(if($cModified -eq $selectedBomModifiedUtc){'match'}else{'MISMATCH'}), schema=$(if($cSchema -eq $CacheSchemaVersion){'match'}else{'MISMATCH'}), parser=$(if($cParser -eq $ParserVersion){'match'}else{'MISMATCH'})"

      if ($fastMatch) {
        # Check all generated outputs exist and are readable
        $outputsOk = $true
        foreach ($p in @($manifestPath, $preparedDataPath, $stagedWorkbookPath)) {
          if (-not (Test-Path -LiteralPath $p)) { $outputsOk = $false; Write-Log "Missing generated output: $p"; break }
          if ((Get-Item -LiteralPath $p).Length -eq 0) { $outputsOk = $false; Write-Log "Empty generated output: $p"; break }
        }
        if ($outputsOk) {
          # Fast cache fresh — reuse
          $sourceHash = $cHash
          $rebuildNeeded = $false
          $cacheStatus = 'fresh'
          Write-Log "BOM cache is fresh. Reusing prepared BOM data."
        } else {
          $rebuildReason = 'missing or empty generated output'
        }
      } else {
        if ($cPath -ne $selectedBom.FullName) { $rebuildReason = 'source path changed' }
        elseif ($cLength -ne $selectedBomSize.ToString()) { $rebuildReason = 'source size changed' }
        elseif ($cModified -ne $selectedBomModifiedUtc) { $rebuildReason = 'source modified time changed' }
        elseif ($cSchema -ne $CacheSchemaVersion) { $rebuildReason = 'cache schema version changed' }
        elseif ($cParser -ne $ParserVersion) { $rebuildReason = 'parser version changed' }
        Write-Log "Fast check failed: $rebuildReason"
      }
    } else {
      $rebuildReason = 'no BomWorkbook section in cache state'
    }
  } else {
    if ($forceRebuild) { $rebuildReason = '--force-rebuild specified' }
    else { $rebuildReason = 'cache state file missing' }
  }
} else {
  $rebuildReason = 'no BOM selected'
}

# === Rebuild (if needed) ===
if ($rebuildNeeded -and $selectedBom) {
  Write-Log "Phase: Rebuilding prepared BOM data (reason: $rebuildReason)"
  Write-Log "BOM cache is stale or missing. Rebuilding prepared BOM data."

  $rebuildStart = Get-Date
  $rebuildFailed = $false
  $newHash = ''
  $newRows = @()
  $newManifestJson = ''
  $newDataJson = ''
  $totalRecords = 0

  try {
    # Calculate SHA-256
    Write-Log "Calculating SHA-256 hash..."
    $newHash = (Get-FileHash -LiteralPath $selectedBom.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    Write-Log "SHA-256: $newHash"

    # Parse BOM workbook
    Write-Log "Parsing BOM workbook (parser v$ParserVersion)..."
    $sheets = @([WvnXlsxReader]::Read($selectedBom.FullName, '.*', 0))

    $allRows = [System.Collections.ArrayList]::new()
    $seenKeys = @{}

    foreach ($sheet in $sheets) {
      Write-Log "  Scanning sheet '$($sheet.Name)' ($($sheet.Rows.Count) rows)..."

      $headerInfo = Get-HeaderInfo $sheet @(@('fg','fg id','finished good','top item id'), @('item number','item no','component','child item id'), @('product name','product','description','item description','child item name'))
      if (-not $headerInfo) { continue }

      $fgCol       = Find-HeaderColumn $headerInfo @('fg','fg id','finished good','top item id')
      $fgNameCol   = Find-HeaderColumn $headerInfo @('fg name','finished good name','top item name')
      $itemCol     = Find-HeaderColumn $headerInfo @('item number','item no','component','child item id')
      $productCol  = Find-HeaderColumn $headerInfo @('product name','product','description','item description','child item name')
      $typeCol     = Find-HeaderColumn $headerInfo @('type','item type','item group','child item group')
      $vendorCol   = Find-HeaderColumn $headerInfo @('vendor','supplier','vendor name')
      $normCol     = Find-HeaderColumn $headerInfo @('norm','total','quantity','qty','per series')
      $buyerCol    = Find-HeaderColumn $headerInfo @('buyer grp','buyer group','purchasing group')
      $levelCol    = Find-HeaderColumn $headerInfo @('bomlevel','bom level','level','component level')

      for ($r = $headerInfo.Row + 1; $r -lt $sheet.Rows.Count; $r++) {
        $row = $sheet.Rows[$r]

        $fg = Normalize-ItemId (Get-RowCell $row $fgCol)
        $itemNumber = Normalize-ItemId (Get-RowCell $row $itemCol)
        $productName = Get-RowCell $row $productCol
        if (-not $fg -or -not $itemNumber -or -not $productName) { continue }

        # Level 1 filter
        if ($levelCol -ge 0) {
          $levelVal = Get-RowCell $row $levelCol
          if ($levelVal -and $levelVal -ne '1') { continue }
        }

        $fgName = if ($fgNameCol -ge 0) { Get-RowCell $row $fgNameCol } else { '' }
        $type = if ($typeCol -ge 0) { Get-RowCell $row $typeCol } else { '' }
        $vendor = if ($vendorCol -ge 0) { Get-RowCell $row $vendorCol } else { '' }
        $norm = if ($normCol -ge 0) { Get-RowCell $row $normCol } else { '' }
        $buyerGroup = if ($buyerCol -ge 0) { Get-RowCell $row $buyerCol } else { '' }

        # Deduplicate: fg::itemNumber::productName, keep row with more populated fields
        $key = "$fg::$itemNumber::$productName".ToLowerInvariant()
        $quality = [int]($fgName -ne '') + [int]($type -ne '') + [int]($vendor -ne '') + [int]($norm -ne '') + [int]($buyerGroup -ne '')

        if ($seenKeys.ContainsKey($key)) {
          if ($quality -gt $seenKeys[$key].Quality) {
            $seenKeys[$key] = @{ Quality=$quality; Index=$allRows.Count }
            [void]$allRows.Add(@{ fg=$fg; fgName=$fgName; itemNumber=$itemNumber; productName=$productName; type=$type; vendor=$vendor; norm=$norm; buyerGroup=$buyerGroup })
          }
        } else {
          $seenKeys[$key] = @{ Quality=$quality; Index=$allRows.Count }
          [void]$allRows.Add(@{ fg=$fg; fgName=$fgName; itemNumber=$itemNumber; productName=$productName; type=$type; vendor=$vendor; norm=$norm; buyerGroup=$buyerGroup })
        }
      }
    }

    $totalRecords = $allRows.Count
    Write-Log "Parsed $totalRecords deduplicated BOM rows from $($sheets.Count) sheets."

    # Build prepared data JSON
    $dataObj = [pscustomobject]@{
      schemaVersion = $CacheSchemaVersion
      parserVersion = $ParserVersion
      sourceName    = $selectedBomName
      sourceHash    = $newHash
      sourceModifiedUtc = $selectedBomModifiedUtc
      sourceSizeBytes   = $selectedBomSize
      generatedUtc  = (Get-Date).ToUniversalTime().ToString('o')
      cacheStatus   = 'rebuilt'
      recordCount   = $totalRecords
      rows          = @($allRows)
    }
    $newDataJson = ConvertTo-Json -Compress -Depth 4 -InputObject $dataObj

    # Build manifest JSON
    $manifestObj = [pscustomobject]@{
      schemaVersion    = $CacheSchemaVersion
      generatedUtc     = (Get-Date).ToUniversalTime().ToString('o')
      cacheStatus      = 'rebuilt'
      source           = [pscustomobject]@{
        name           = $selectedBomName
        modifiedUtc    = $selectedBomModifiedUtc
        sizeBytes      = $selectedBomSize
        hash           = $newHash
      }
      preparedDataPath = 'system/generated/inspection_bom_data.json'
      stagedWorkbookPath = 'system/generated/inspection_bom_source.xlsx'
      recordCount      = $totalRecords
    }
    $newManifestJson = ConvertTo-Json -Compress -Depth 4 -InputObject $manifestObj

    # Write prepared data (atomic)
    Write-Log "Writing prepared data JSON ($totalRecords records)..."
    Write-AtomicJson -FilePath $preparedDataPath -JsonContent $newDataJson
    Write-Log "Prepared data written: $preparedDataPath"

    # Stage workbook (atomic via temp copy)
    $tmpWorkbookPath = "$stagedWorkbookPath.tmp"
    Copy-Item -LiteralPath $selectedBom.FullName -Destination $tmpWorkbookPath -Force
    if (Test-Path -LiteralPath $tmpWorkbookPath) {
      Move-Item -LiteralPath $tmpWorkbookPath -Destination $stagedWorkbookPath -Force
      Write-Log "Staged workbook: $stagedWorkbookPath"
    }

    # Write manifest (atomic)
    Write-AtomicJson -FilePath $manifestPath -JsonContent $newManifestJson
    Write-Log "Manifest written: $manifestPath"

    # Also write legacy manifest at root for backward compatibility
    $legacyManifest = ConvertTo-Json -Compress -Depth 3 -InputObject @(
      [pscustomobject]@{ path='system/generated/inspection_bom_source.xlsx'; name=$selectedBomName; modifiedUtc=$selectedBomModifiedUtc }
    )
    Write-AtomicJson -FilePath $legacyManifestPath -JsonContent $legacyManifest
    Write-Log "Legacy manifest written: $legacyManifestPath"

    # Write cache state (atomic)
    $cacheLines = @(
      "[BomWorkbook]",
      "SourceFullPath=$($selectedBom.FullName)",
      "SourceFileName=$selectedBomName",
      "SourceLength=$selectedBomSize",
      "LastWriteTimeUtc=$selectedBomModifiedUtc",
      "Hash=$newHash",
      "CacheSchemaVersion=$CacheSchemaVersion",
      "ParserVersion=$ParserVersion",
      "",
      "[GeneratedOutput]",
      "ManifestPath=$manifestPath",
      "PreparedDataPath=$preparedDataPath",
      "StagedWorkbookPath=$stagedWorkbookPath",
      "GeneratedUtc=$((Get-Date).ToUniversalTime().ToString('o'))"
    )
    Write-AtomicIni -FilePath $cacheStatePath -Lines $cacheLines
    Write-Log "Cache state written: $cacheStatePath"

    $sourceHash = $newHash
    $cacheStatus = 'rebuilt'

    $rebuildDuration = [math]::Round(((Get-Date) - $rebuildStart).TotalSeconds, 1)
    Write-Log "Rebuild completed in ${rebuildDuration}s."

  } catch {
    Write-Log "ERROR during rebuild: $_"
    $rebuildFailed = $true

    # Clean up temp files
    @("$preparedDataPath.tmp", "$manifestPath.tmp", "$stagedWorkbookPath.tmp") | ForEach-Object {
      if (Test-Path -LiteralPath $_) { Remove-Item -LiteralPath $_ -Force -ErrorAction SilentlyContinue }
    }

    # Fall back to previous cache if available
    if ((Test-Path -LiteralPath $preparedDataPath) -and (Test-Path -LiteralPath $manifestPath) -and (Test-Path -LiteralPath $stagedWorkbookPath)) {
      Write-Log "Previous cache preserved. Using stale fallback."
      $cacheStatus = 'stale'
      try {
        $prevCacheState = Read-IniFile -FilePath $cacheStatePath
        if ($prevCacheState.ContainsKey('BomWorkbook')) {
          $sourceHash = $prevCacheState['BomWorkbook']['Hash']
          $selectedBomName = $prevCacheState['BomWorkbook']['SourceFileName']
          $selectedBomModifiedUtc = $prevCacheState['BomWorkbook']['LastWriteTimeUtc']
        }
      } catch {}
    } else {
      Write-Log "No valid previous cache available."
      $cacheStatus = 'unavailable'
    }
  }

  # Release memory
  $allRows = $null
  $seenKeys = $null
  $sheets = $null
  [System.GC]::Collect()

} elseif ($rebuildNeeded -and -not $selectedBom) {
  Write-Log "No BOM selected. Generating empty manifest."
  $cacheStatus = 'unavailable'

  $emptyManifest = '{ "schemaVersion": "' + $CacheSchemaVersion + '", "available": false, "cacheStatus": "unavailable", "message": "No valid Weekly BOM workbook was found." }'
  Set-Content -Encoding UTF8 -LiteralPath $manifestPath -Value $emptyManifest
  Set-Content -Encoding UTF8 -LiteralPath $legacyManifestPath -Value '[]'
} else {
  # Cache is fresh — reuse existing; just update legacy manifest
  Write-Log "Cache is fresh. Skipping all preparation."

  if (Test-Path -LiteralPath $manifestPath) {
    try { $manifestData = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json } catch {}
  }

  # Rewrite legacy manifest to ensure it points to current staged file
  if (Test-Path -LiteralPath $stagedWorkbookPath) {
    $legacyManifest = ConvertTo-Json -Compress -Depth 3 -InputObject @(
      [pscustomobject]@{ path='system/generated/inspection_bom_source.xlsx'; name=$selectedBomName; modifiedUtc=$selectedBomModifiedUtc }
    )
    Write-AtomicJson -FilePath $legacyManifestPath -JsonContent $legacyManifest
  }
}

# === Master-carton picture index and browser cache ===
Write-Log "Phase: Indexing master-carton pictures"
$pictureRecords = [System.Collections.ArrayList]::new()
$pictureLibraryAvailable = $false
$pictureLibraryMessage = ''
$previousPicturesByName = @{}

if (Test-Path -LiteralPath $pictureIndexPath) {
  try {
    $previousIndex = Get-Content -Raw -LiteralPath $pictureIndexPath | ConvertFrom-Json
    foreach ($previousPicture in @($previousIndex.pictures)) {
      if ($previousPicture.filename) {
        $previousPicturesByName[[string]$previousPicture.filename] = $previousPicture
      }
    }
  } catch {
    Write-Log "WARNING: Previous picture index could not be read; hashes will be rebuilt."
  }
}

if ($resolvedMasterCartonPictureDir -and
    (Test-Path -LiteralPath $resolvedMasterCartonPictureDir -PathType Container)) {
  $pictureLibraryAvailable = $true
  $pictureFiles = @(
    Get-ChildItem -LiteralPath $resolvedMasterCartonPictureDir -File -ErrorAction SilentlyContinue |
      Where-Object {
        $_.Extension.ToLowerInvariant() -in @('.png', '.jpg', '.jpeg', '.webp')
      } |
      Sort-Object Name
  )

  foreach ($pictureFile in $pictureFiles) {
    if ($pictureFile.Name -notmatch '^(.+?)_image(\d+)(?:_(\d+))?\.(png|jpg|jpeg|webp)$') {
      Write-Log "  Skipped unrecognized picture filename: $($pictureFile.Name)"
      continue
    }

    $cartonId = Normalize-ItemId $matches[1]
    $sequence = [int]$matches[2]
    $variant = if ($matches[3]) { [int]$matches[3] } else { $null }
    $extension = $matches[4].ToLowerInvariant()
    $sourceModifiedUtc = $pictureFile.LastWriteTimeUtc.ToString('o')
    $previousPicture = $previousPicturesByName[$pictureFile.Name]
    $sourceHash = $null

    if (-not $forceRebuild -and
        $previousPicture -and
        [int64]$previousPicture.sizeBytes -eq $pictureFile.Length -and
        [string]$previousPicture.modifiedUtc -eq $sourceModifiedUtc -and
        $previousPicture.sha256) {
      $sourceHash = [string]$previousPicture.sha256
    } else {
      $sourceHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $pictureFile.FullName).Hash
    }

    $cachedPath = Join-Path $pictureCacheDir $pictureFile.Name
    $copyReason = $null
    if ($forceRebuild) {
      $copyReason = 'force rebuild'
    } elseif (-not (Test-Path -LiteralPath $cachedPath -PathType Leaf)) {
      $copyReason = 'missing'
    } else {
      $cachedFile = Get-Item -LiteralPath $cachedPath
      if ($cachedFile.Length -ne $pictureFile.Length) {
        $copyReason = 'size changed'
      } elseif ($pictureFile.LastWriteTimeUtc -gt $cachedFile.LastWriteTimeUtc) {
        $copyReason = 'source newer'
      } else {
        $cachedHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $cachedPath).Hash
        if ($cachedHash -ne $sourceHash) {
          $copyReason = 'hash changed'
        }
      }
    }

    if ($copyReason) {
      Copy-Item -LiteralPath $pictureFile.FullName -Destination $cachedPath -Force
      (Get-Item -LiteralPath $cachedPath).LastWriteTimeUtc = $pictureFile.LastWriteTimeUtc
      Write-Log "  Cached $($pictureFile.Name) ($copyReason)"
    }

    [void]$pictureRecords.Add([pscustomobject]@{
      cartonId = $cartonId
      filename = $pictureFile.Name
      sequence = $sequence
      variant = $variant
      extension = $extension
      relativePath = ('system/cache/master-carton-pictures/' + $pictureFile.Name)
      modifiedUtc = $sourceModifiedUtc
      sizeBytes = [int64]$pictureFile.Length
      sha256 = $sourceHash
      exactDuplicateCount = 0
    })
  }

  foreach ($hashGroup in @($pictureRecords | Group-Object sha256)) {
    $duplicateCount = [math]::Max(0, $hashGroup.Count - 1)
    foreach ($record in $hashGroup.Group) {
      $record.exactDuplicateCount = $duplicateCount
    }
  }

  $pictureRecords = [System.Collections.ArrayList]@(
    $pictureRecords |
      Sort-Object `
        @{ Expression = { $_.cartonId }; Ascending = $true },
        @{ Expression = { $_.sequence }; Ascending = $true },
        @{ Expression = { if ($null -eq $_.variant) { -1 } else { $_.variant } }; Ascending = $true },
        @{ Expression = { $_.filename }; Ascending = $true }
  )
  $pictureLibraryMessage = "Indexed $($pictureRecords.Count) picture(s)."
} else {
  $pictureLibraryMessage = 'Master-carton picture folder is unavailable. Manual image mode remains available.'
  Write-Log "WARNING: $pictureLibraryMessage"
}

$cartonIdCount = @($pictureRecords.cartonId | Sort-Object -Unique).Count
$pictureIndexObject = [pscustomobject]@{
  schemaVersion = 1
  generatedUtc = (Get-Date).ToUniversalTime().ToString('o')
  available = $pictureLibraryAvailable
  sourceFolderName = $masterCartonPictureFolderName
  message = $pictureLibraryMessage
  recordCount = $pictureRecords.Count
  cartonIdCount = $cartonIdCount
  pictures = @($pictureRecords)
}
Write-AtomicJson -FilePath $pictureIndexPath -JsonContent (
  ConvertTo-Json -Compress -Depth 6 -InputObject $pictureIndexObject
)
Write-Log "Picture index: $pictureIndexPath"
Write-Log "Picture records: $($pictureRecords.Count); carton IDs: $cartonIdCount"

# === Launch summary ===
Write-Log "--- Launch Summary ---"
Write-Log "  BOM search directory: $(if($resolvedWeeklyBomDir){$resolvedWeeklyBomDir}else{'[None]'})"
Write-Log "  BOM candidates found: $bomFilesFound"
Write-Log "  Selected BOM: $(if($selectedBomName){$selectedBomName}else{'[None]'})"
Write-Log "  Source modified UTC: $(if($selectedBomModifiedUtc){$selectedBomModifiedUtc}else{'[N/A]'})"
Write-Log "  Source size: $(if($selectedBomSize){$selectedBomSize}else{0}) bytes"
Write-Log "  SHA-256 calculated: $(if($rebuildNeeded -and $selectedBom){'yes'}else{'no (fast check reused)'})"
Write-Log "  Cache status: $cacheStatus"
if ($rebuildReason) { Write-Log "  Rebuild reason: $rebuildReason" }
Write-Log "  Manifest path: $manifestPath"
Write-Log "  Prepared data: $preparedDataPath"
Write-Log "  Staged workbook: $stagedWorkbookPath"
Write-Log "  Picture library available: $pictureLibraryAvailable"
Write-Log "  Picture records: $($pictureRecords.Count)"
Write-Log "  Picture carton IDs: $cartonIdCount"
$totalDuration = [math]::Round(((Get-Date) - $script:logStart).TotalSeconds, 1)
Write-Log "  Total launcher duration: ${totalDuration}s"

# === Detect browser ===
$browser = $null
$browserName = ''
$chromePaths = @(
  (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),
  (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'),
  (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
)
foreach ($cp in $chromePaths) {
  if (Test-Path -LiteralPath $cp) { $browser = $cp; $browserName = 'Chrome'; break }
}
if (-not $browser) {
  $edgePaths = @(
    (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
    (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
  )
  foreach ($ep in $edgePaths) {
    if (Test-Path -LiteralPath $ep) { $browser = $ep; $browserName = 'Edge'; break }
  }
}

Write-Log "Browser: $(if($browser){"$browserName at $browser"}else{'[Not found]'})"

if (-not $browser) {
  Write-Log "ERROR: Chrome and Edge not found in standard locations."
  Write-Log "  Checked: $($chromePaths -join ', ')"

  $showConsole = $true
  if ($config.ContainsKey('Launcher') -and $config['Launcher'].ContainsKey('ShowConsole')) {
    $showConsole = $config['Launcher']['ShowConsole'] -eq 'true'
  }
  if ($showConsole) {
    Write-Host ""
    Write-Host "ERROR: Chrome and Edge not found." -ForegroundColor Red
    Write-Host "Please install Chrome or open Inspection Standard Tool.html manually." -ForegroundColor Red
    Write-Host ""
  }
  Write-Log "Done (browser not found)."
  exit 1
}

# === Launch browser ===
if ($prepareOnly) {
  Write-Log "Data prepared. Skipping browser launch (--prepare-only)."
} else {
  $htmlPath = Join-Path $AppRoot 'Inspection Standard Tool.html'
  Write-Log "Launching browser: $browser"
  Write-Log "HTML path: $htmlPath"
  Write-Log "Chrome profile: $chromeProfileDir"

  $showConsole = $true
  if ($config.ContainsKey('Launcher') -and $config['Launcher'].ContainsKey('ShowConsole')) {
    $showConsole = $config['Launcher']['ShowConsole'] -eq 'true'
  }

  Start-Process -FilePath $browser -ArgumentList "--user-data-dir=`"$chromeProfileDir`"", '--allow-file-access-from-files', "`"$htmlPath`""

  if ($showConsole) {
    Write-Host ""
    Write-Host "Inspection Standard Tool launched. You may close this window." -ForegroundColor Green
    Write-Host ""
  } else {
    # Hide PowerShell console when not needed
    exit 0
  }
}

Write-Log "Done."
exit 0
