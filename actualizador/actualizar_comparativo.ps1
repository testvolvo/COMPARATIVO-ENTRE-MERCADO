param(
  [switch]$Configurar,
  [switch]$SinPublicar
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$RepoRoot = Split-Path -Parent $PSScriptRoot
$ConfigDir = Join-Path $env:LOCALAPPDATA "ComparativoModelosVolvo"
$ConfigPath = Join-Path $ConfigDir "configuracion.json"
$ColumnsPath = Join-Path $RepoRoot "columnas_publicas.json"
$ModelsPath = Join-Path $RepoRoot "modelos.json"
$VersionPath = Join-Path $RepoRoot "version.json"
$SheetDefault = "BBDD MARCAS"

function Write-Utf8NoBom([string]$Path,[string]$Text) {
  $enc = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($Path,$Text,$enc)
}

function Select-ExcelFile {
  Add-Type -AssemblyName System.Windows.Forms
  $dialog = New-Object System.Windows.Forms.OpenFileDialog
  $dialog.Title = "Selecciona el Excel sincronizado desde SharePoint"
  $dialog.Filter = "Archivos Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm"
  $dialog.Multiselect = $false
  if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    throw "No se seleccionó el Excel."
  }
  return $dialog.FileName
}

function Save-Configuration([string]$ExcelPath,[string]$SheetName) {
  New-Item -ItemType Directory -Path $ConfigDir -Force | Out-Null
  $config = [ordered]@{
    excelPath = $ExcelPath
    sheetName = $SheetName
    autoPush = $true
  }
  Write-Utf8NoBom $ConfigPath ($config | ConvertTo-Json -Depth 5)
  Write-Host "Configuración guardada en: $ConfigPath" -ForegroundColor Green
}

function Load-Or-CreateConfiguration {
  if ($Configurar -or -not (Test-Path $ConfigPath)) {
    $path = Select-ExcelFile
    Save-Configuration $path $SheetDefault
  }
  $cfg = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
  if (-not (Test-Path -LiteralPath $cfg.excelPath)) {
    Write-Host "La ruta guardada ya no existe. Selecciona nuevamente el archivo." -ForegroundColor Yellow
    $path = Select-ExcelFile
    $sheetToSave = if ($cfg.sheetName) { [string]$cfg.sheetName } else { $SheetDefault }
    Save-Configuration $path $sheetToSave
    $cfg = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
  }
  return $cfg
}

function Find-Git {
  $cmd = Get-Command git.exe -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }
  $desktopRoot = Join-Path $env:LOCALAPPDATA "GitHubDesktop"
  if (Test-Path $desktopRoot) {
    $candidate = Get-ChildItem -Path $desktopRoot -Filter git.exe -Recurse -ErrorAction SilentlyContinue |
      Where-Object { $_.FullName -match "\\resources\\app\\git\\cmd\\git.exe$" } |
      Sort-Object FullName -Descending | Select-Object -First 1
    if ($candidate) { return $candidate.FullName }
  }
  return $null
}

function Convert-ExcelValue($Value) {
  if ($null -eq $Value) { return $null }
  if ($Value -is [string]) {
    $trimmed = $Value.Trim()
    if ($trimmed.Length -eq 0) { return $null }
    return $trimmed
  }
  if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [int]) { return $Value }
  if ($Value -is [bool]) { return $Value }
  return [string]$Value
}

function Export-ModelsFromExcel([string]$ExcelPath,[string]$SheetName,[string[]]$PublicColumns) {
  $excel = $null; $book = $null; $sheet = $null; $used = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $book = $excel.Workbooks.Open($ExcelPath,0,$true)
    try { $sheet = $book.Worksheets.Item($SheetName) }
    catch { throw "No se encontró la hoja '$SheetName' en el Excel." }
    $used = $sheet.UsedRange
    $values = $used.Value2
    $rowCount = $used.Rows.Count
    $colCount = $used.Columns.Count
    if ($rowCount -lt 2) { throw "La hoja '$SheetName' no contiene filas de datos." }

    $headerIndex = @{}
    for ($c=1; $c -le $colCount; $c++) {
      $header = [string]$values[1,$c]
      if (-not [string]::IsNullOrWhiteSpace($header)) { $headerIndex[$header.Trim()] = $c }
    }
    $missing = @($PublicColumns | Where-Object { -not $headerIndex.ContainsKey($_) })
    if ($missing.Count -gt 0) {
      Write-Host "Columnas no encontradas (se dejarán vacías): $($missing -join ', ')" -ForegroundColor Yellow
    }

    $models = New-Object System.Collections.Generic.List[object]
    for ($r=2; $r -le $rowCount; $r++) {
      $record = [ordered]@{}
      $hasIdentity = $false
      foreach ($column in $PublicColumns) {
        $value = $null
        if ($headerIndex.ContainsKey($column)) { $value = Convert-ExcelValue $values[$r,$headerIndex[$column]] }
        $record[$column] = $value
        if (($column -eq "ID" -or $column -eq "DESIGNACION DE VENTA" -or $column -eq "MODELO") -and $null -ne $value) { $hasIdentity = $true }
      }
      if (-not $hasIdentity) { continue }
      $record["_uid"] = $models.Count + 1
      $models.Add([pscustomobject]$record)
    }
    return ,$models.ToArray()
  }
  finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($obj in @($used,$sheet,$book,$excel)) {
      if ($obj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Publish-With-Git([string]$GitPath) {
  & $GitPath -C $RepoRoot add -- modelos.json version.json columnas_publicas.json index.html
  & $GitPath -C $RepoRoot diff --cached --quiet
  if ($LASTEXITCODE -eq 0) {
    Write-Host "No hay cambios nuevos para publicar." -ForegroundColor Cyan
    return
  }
  $stamp = Get-Date -Format "yyyy-MM-dd HH:mm"
  & $GitPath -C $RepoRoot commit -m "Actualizar base comparativo $stamp"
  if ($LASTEXITCODE -ne 0) { throw "Git no pudo crear el commit." }
  & $GitPath -C $RepoRoot push
  if ($LASTEXITCODE -ne 0) { throw "Los archivos se actualizaron, pero Git no pudo hacer push. Abre GitHub Desktop y presiona Push origin." }
  Write-Host "Cambios publicados en GitHub Pages." -ForegroundColor Green
}

try {
  Write-Host ""; Write-Host "=== ACTUALIZADOR COMPARATIVO DE MODELOS ===" -ForegroundColor Cyan
  $cfg = Load-Or-CreateConfiguration
  if ($Configurar) {
    Write-Host "Configuración terminada. Ejecuta Actualizar_Comparativo.bat para publicar." -ForegroundColor Green
    exit 0
  }
  if (-not (Test-Path $ColumnsPath)) { throw "Falta columnas_publicas.json en la raíz del repositorio." }
  $columns = @(Get-Content -LiteralPath $ColumnsPath -Raw -Encoding UTF8 | ConvertFrom-Json)
  if ($columns.Count -eq 0) { throw "columnas_publicas.json está vacío." }

  Write-Host "Excel: $($cfg.excelPath)"
  Write-Host "Hoja:  $($cfg.sheetName)"
  Write-Host "Leyendo la base..." -ForegroundColor Cyan
  $models = Export-ModelsFromExcel $cfg.excelPath $cfg.sheetName $columns
  $source = Get-Item -LiteralPath $cfg.excelPath
  $now = (Get-Date).ToUniversalTime().ToString("o")
  $sourceModified = $source.LastWriteTimeUtc.ToString("o")

  $baseMeta = [ordered]@{
    schemaVersion = 1
    updatedAt = $now
    sourceModified = $sourceModified
    count = $models.Count
    sheet = $cfg.sheetName
  }
  $payloadForHash = [ordered]@{ meta = $baseMeta; models = $models }
  $compactForHash = $payloadForHash | ConvertTo-Json -Depth 100 -Compress
  $sha = [System.Security.Cryptography.SHA256]::Create()
  try { $hashBytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($compactForHash)) }
  finally { $sha.Dispose() }
  $hash = -join ($hashBytes | ForEach-Object { $_.ToString("x2") })
  $baseMeta["hash"] = $hash
  $payload = [ordered]@{ meta = $baseMeta; models = $models }
  Write-Utf8NoBom $ModelsPath ($payload | ConvertTo-Json -Depth 100 -Compress)
  Write-Utf8NoBom $VersionPath ($baseMeta | ConvertTo-Json -Depth 10)
  Write-Host "$($models.Count) configuraciones exportadas." -ForegroundColor Green

  if ($SinPublicar -or -not $cfg.autoPush) {
    Write-Host "Archivos actualizados localmente. No se realizó push." -ForegroundColor Yellow
    exit 0
  }
  $git = Find-Git
  if (-not $git) {
    Write-Host "No se encontró Git. Los archivos quedaron actualizados." -ForegroundColor Yellow
    Write-Host "Abre GitHub Desktop, crea el commit y presiona Push origin."
    exit 0
  }
  Publish-With-Git $git
}
catch {
  Write-Host ""; Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host "No se modificó el Excel original." -ForegroundColor Yellow
  exit 1
}
