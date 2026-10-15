param(
  [string]$ConfigPath = "$PSScriptRoot\..\references\pipeline_config.example.json"
)

$ErrorActionPreference = "Stop"

function Resolve-ConfiguredPath {
  param(
    [AllowNull()][object]$Value,
    [string]$BasePath
  )
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace("$Value")) {
    return $null
  }
  $candidate = "$Value"
  if ([System.IO.Path]::IsPathRooted($candidate)) {
    return $candidate
  }
  return [System.IO.Path]::GetFullPath((Join-Path $BasePath $candidate))
}

function Resolve-Executable {
  param(
    [AllowNull()][object]$ExplicitPath,
    [string[]]$CommandNames,
    [string]$Label,
    [bool]$Required = $true
  )
  if ($null -ne $ExplicitPath -and -not [string]::IsNullOrWhiteSpace("$ExplicitPath")) {
    $resolved = "$ExplicitPath"
    if (-not (Test-Path -LiteralPath $resolved)) {
      throw "$Label was configured but not found: $resolved"
    }
    return (Resolve-Path -LiteralPath $resolved).Path
  }
  foreach ($name in $CommandNames) {
    $command = Get-Command $name -ErrorAction SilentlyContinue
    if ($null -ne $command) {
      return $command.Source
    }
  }
  if ($Required) {
    throw "$Label was not found. Add it to PATH or set runtime.$($Label.ToLower()) in the config."
  }
  return $null
}

$resolvedConfig = (Resolve-Path -LiteralPath $ConfigPath).Path
$config = Get-Content -LiteralPath $resolvedConfig -Raw -Encoding UTF8 | ConvertFrom-Json
$skillRoot = Split-Path -Parent $PSScriptRoot
$configRoot = Split-Path -Parent $resolvedConfig
$runRoot = (Get-Location).Path

$outputRoot = if ([System.IO.Path]::IsPathRooted("$($config.output_dir)")) {
  "$($config.output_dir)"
} else {
  [System.IO.Path]::GetFullPath((Join-Path $runRoot "$($config.output_dir)"))
}
$null = New-Item -ItemType Directory -Force -Path $outputRoot

$python = Resolve-Executable -ExplicitPath $config.runtime.python -CommandNames @("python", "python3") -Label "Python"
$pdfScript = Join-Path $PSScriptRoot "run_pdf_json.py"
$policyValue = "$($config.page_policy)"
if ([System.IO.Path]::IsPathRooted($policyValue)) {
  $policy = $policyValue
} elseif (Test-Path -LiteralPath (Join-Path $configRoot $policyValue)) {
  $policy = [System.IO.Path]::GetFullPath((Join-Path $configRoot $policyValue))
} else {
  $policy = [System.IO.Path]::GetFullPath((Join-Path $skillRoot $policyValue))
}

$structurePdf = Resolve-ConfiguredPath -Value $config.inputs.structure_pdf -BasePath $configRoot
if ($null -eq $structurePdf -or -not (Test-Path -LiteralPath $structurePdf)) {
  throw "inputs.structure_pdf must point to an existing PDF: $structurePdf"
}

$pdftoppmExplicit = Resolve-ConfiguredPath -Value $config.render.pdftoppm -BasePath $configRoot
$pdftoppm = Resolve-Executable -ExplicitPath $pdftoppmExplicit -CommandNames @("pdftoppm") -Label "Pdftoppm" -Required $false
$pdfArgs = @(
  $pdfScript,
  "--pdf", $structurePdf,
  "--policy", $policy,
  "--output", (Join-Path $outputRoot "structure"),
  "--deep-dpi", "$($config.render.deep_dpi)"
)
if ($null -ne $pdftoppm) {
  $pdfArgs += @("--pdftoppm", $pdftoppm)
}

& $python @pdfArgs
if ($LASTEXITCODE -ne 0) { throw "structure_pdf_to_json failed with exit code $LASTEXITCODE" }

$stages = [System.Collections.Generic.List[string]]::new()
$stages.Add("structure_pdf_to_json")

if ($null -ne $config.inputs.email_pdf -and -not [string]::IsNullOrWhiteSpace("$($config.inputs.email_pdf)")) {
  $emailPdf = Resolve-ConfiguredPath -Value $config.inputs.email_pdf -BasePath $configRoot
  $emailArgs = @(
    $pdfScript,
    "--pdf", $emailPdf,
    "--output", (Join-Path $outputRoot "email"),
    "--deep-dpi", "$($config.render.deep_dpi)",
    "--default-mode", "deep_processing"
  )
  if ($null -ne $pdftoppm) {
    $emailArgs += @("--pdftoppm", $pdftoppm)
  }
  & $python @emailArgs
  if ($LASTEXITCODE -ne 0) { throw "email_pdf_to_json failed with exit code $LASTEXITCODE" }
  $stages.Add("email_pdf_to_json")
}

if ($null -ne $config.inputs.excel_template -and -not [string]::IsNullOrWhiteSpace("$($config.inputs.excel_template)")) {
  $excelTemplate = Resolve-ConfiguredPath -Value $config.inputs.excel_template -BasePath $configRoot
  $node = Resolve-Executable -ExplicitPath $config.runtime.node -CommandNames @("node") -Label "Node"
  $nodeModules = Resolve-ConfiguredPath -Value $config.runtime.node_modules -BasePath $configRoot
  $skillNodeModules = Join-Path $skillRoot "node_modules"
  if ($null -ne $nodeModules -and -not (Test-Path -LiteralPath $skillNodeModules)) {
    New-Item -ItemType Junction -Path $skillNodeModules -Target $nodeModules | Out-Null
  }
  & $node (Join-Path $PSScriptRoot "read_template.mjs") $excelTemplate (Join-Path $outputRoot "template_schema.json")
  if ($LASTEXITCODE -ne 0) {
    throw "excel_template_to_schema failed. Confirm Node.js and @oai/artifact-tool are available, or leave inputs.excel_template null."
  }
  $stages.Add("excel_template_to_schema")
}

if ($null -ne $config.excel_mapping_config -and -not [string]::IsNullOrWhiteSpace("$($config.excel_mapping_config)")) {
  $mappingConfigPath = Resolve-ConfiguredPath -Value $config.excel_mapping_config -BasePath $configRoot
  & $python (Join-Path $PSScriptRoot "write_excel_from_mapping.py") --config $mappingConfigPath
  if ($LASTEXITCODE -ne 0) { throw "excel_mapping failed with exit code $LASTEXITCODE" }
  $stages.Add("excel_mapping")
}

$runLog = [ordered]@{
  schema = "dbh-formatting/run-log@0.1"
  generated_at = (Get-Date).ToUniversalTime().ToString("o")
  project_code = $config.project_code
  output_dir = $outputRoot
  stages = $stages
  note = "estimated_input_tokens is a local estimate, not an API billing value. Email content is processed only when inputs.email_pdf is explicitly configured."
}
$runLog | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $outputRoot "run_log.json") -Encoding UTF8
Write-Output "DBH-Formatting completed: $outputRoot"

