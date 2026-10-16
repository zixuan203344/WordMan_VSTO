param(
    [string]$MSBuildPath = 'D:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe',
    [switch]$WordIntegration
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$buildDir = Join-Path ([System.IO.Path]::GetTempPath()) ('WordManStyleAudit-' + [guid]::NewGuid().ToString('N'))
$compiler = Join-Path (Split-Path -Parent $MSBuildPath) 'Roslyn\csc.exe'
$wordInterop = Join-Path $env:windir 'assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll'
if (!(Test-Path -LiteralPath $MSBuildPath) -or !(Test-Path -LiteralPath $compiler)) {
    throw 'Pass -MSBuildPath with a Visual Studio installation that supports Office/VSTO projects.'
}
New-Item -ItemType Directory -Path $buildDir | Out-Null
$previousConfig = $env:ACC_PRODUCT_CONFIG_V3
try {
    $env:ACC_PRODUCT_CONFIG_V3 = ''
    # Compile without signing, registering, or deploying the installed add-in.
    & $MSBuildPath (Join-Path $repoRoot 'WordMan.csproj') /t:Compile /p:Configuration=Debug /p:SignManifests=false "/p:OutputPath=$buildDir\" "/p:IntermediateOutputPath=$buildDir\obj\" /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) { throw 'Project compilation failed.' }
    Copy-Item -LiteralPath (Join-Path $buildDir 'obj\WordMan.dll') -Destination (Join-Path $buildDir 'WordMan.dll')
    $testExe = Join-Path $buildDir 'StyleSettingsRegression.exe'
    & $compiler /nologo /target:exe /platform:x64 "/out:$testExe" "/reference:$buildDir\WordMan.dll" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Core.dll "/link:$wordInterop" (Join-Path $PSScriptRoot 'StyleSettingsRegression.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed.' }
    if ($WordIntegration) { & $testExe --word } else { & $testExe }
    if ($LASTEXITCODE -ne 0) { throw "Regression failed. Build output: $buildDir" }
    Write-Output "Build output: $buildDir"
}
finally { $env:ACC_PRODUCT_CONFIG_V3 = $previousConfig }
