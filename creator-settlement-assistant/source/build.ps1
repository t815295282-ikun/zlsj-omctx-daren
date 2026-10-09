$ErrorActionPreference = 'Stop'
$taskSource = $PSScriptRoot
$taskOutput = Split-Path -Parent $taskSource
$taskCompiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (!(Test-Path -LiteralPath $taskCompiler)) {
    $taskCompiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
$taskArgs = @('/nologo', '/target:winexe', '/platform:anycpu', '/optimize+', '/utf8output',
    ('/out:' + (Join-Path $taskOutput '达人结算智能助手.exe')),
    '/reference:System.dll', '/reference:System.Core.dll', '/reference:System.Data.dll',
    '/reference:System.Data.DataSetExtensions.dll', '/reference:System.Drawing.dll',
    '/reference:System.Windows.Forms.dll', '/reference:System.Xml.dll', '/reference:System.Xml.Linq.dll',
    '/reference:System.IO.Compression.dll', '/reference:System.IO.Compression.FileSystem.dll',
    '/reference:System.Web.Extensions.dll')
foreach ($taskTemplate in @('check', 'payment', 'adjust')) {
    $taskArgs += '/resource:' + (Join-Path $taskSource ('templates\' + $taskTemplate + '.xlsx')) + ',' + $taskTemplate + '.xlsx'
}
$taskArgs += Join-Path $taskSource 'Assistant.cs'
$taskArgs += Join-Path $taskSource 'Window.cs'
$taskArgs += Join-Path $taskSource 'Features.cs'
$taskArgs += Join-Path $taskSource 'Exporting.cs'
$taskArgs += Join-Path $taskSource 'History.cs'
$taskArgs += Join-Path $taskSource 'GridCopy.cs'
$taskArgs += Join-Path $taskSource 'TableEditing.cs'
$taskArgs += Join-Path $taskSource 'Dpi.cs'
$taskArgs += '/win32manifest:' + (Join-Path $taskSource 'app.manifest')
$taskArgs += Join-Path $taskSource 'ExcelEditing.cs'
& $taskCompiler $taskArgs
if ($LASTEXITCODE -ne 0) { throw '编译失败' }
Write-Host '编译完成：达人结算智能助手.exe'

