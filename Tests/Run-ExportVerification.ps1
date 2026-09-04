param([string]$CscPath)

$ErrorActionPreference = 'Stop'
$workspace = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$project = Join-Path $workspace 'EasyAccessRules/EasyAccessRules'

if (-not $CscPath) {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $CscPath = & $vswhere -latest -products '*' -find 'MSBuild\**\Bin\Roslyn\csc.exe' | Select-Object -First 1
}

if (-not $CscPath -or -not (Test-Path -LiteralPath $CscPath)) {
    throw 'Pass -CscPath with the path to the Visual Studio Roslyn csc.exe compiler.'
}

$testId = [Guid]::NewGuid().ToString('N')
$executable = Join-Path ([IO.Path]::GetTempPath()) "EasyAccessRules-$testId.exe"
$export = Join-Path ([IO.Path]::GetTempPath()) "EasyAccessRules-$testId.xml"

try {
    & $CscPath /nologo /langversion:latest /target:exe "/out:$executable" /main:ExportVerification `
        /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Xml.dll /r:System.Core.dll `
        (Join-Path $PSScriptRoot 'ExportVerification.cs') `
        (Join-Path $project 'MainWindow.cs') (Join-Path $project 'MainWindow.Designer.cs') `
        (Join-Path $project 'WordXmlExporter.cs')

    if ($LASTEXITCODE -ne 0) { throw 'Verification build failed.' }

    & $executable (Join-Path $workspace 'Easy Access Rules for Air Operations (XML).xml') $export

    if ($LASTEXITCODE -ne 0) { throw 'Export verification failed.' }
}
finally {
    foreach ($file in @($executable, $export)) {
        if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file }
    }
}
