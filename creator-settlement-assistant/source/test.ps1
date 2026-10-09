param([Parameter(Mandatory=$true)][string]$ResultDirectory)
$ErrorActionPreference='Stop'
if ((Test-Path -LiteralPath $ResultDirectory) -and @(Get-ChildItem -LiteralPath $ResultDirectory -Force).Count -gt 0) { throw '请选择新的空测试结果文件夹；不要使用程序目录或真实收款资料目录。' }
& (Join-Path $PSScriptRoot 'build.ps1')
$taskApp=Join-Path (Split-Path -Parent $PSScriptRoot) '达人结算智能助手.exe'
$taskCompiler=Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (!(Test-Path -LiteralPath $taskCompiler)) { $taskCompiler=Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
New-Item -ItemType Directory -Force -Path $ResultDirectory | Out-Null
Copy-Item -LiteralPath $taskApp -Destination $ResultDirectory
$taskReferences=@('/nologo','/utf8output','/target:exe',('/reference:'+$taskApp),'/reference:System.dll','/reference:System.Core.dll','/reference:System.Data.dll','/reference:System.Data.DataSetExtensions.dll','/reference:System.Drawing.dll','/reference:System.Windows.Forms.dll','/reference:System.Xml.dll','/reference:System.Xml.Linq.dll','/reference:System.IO.Compression.dll','/reference:System.IO.Compression.FileSystem.dll')
foreach($taskPair in @(@('Regression.exe','TestsV13'),@('OpenRegression.exe','TestsOpen'),@('RegressionHistory.exe','TestsHistory'),@('EditingRegression.exe','TestsEditing'),@('AdjustmentRegression.exe','TestsAdjustment'),@('ExcelRegression.exe','TestsExcelEditing'))){
 $taskSources=@((Join-Path $PSScriptRoot 'TestsV13.cs'))
 if($taskPair[1] -ne 'TestsV13'){$taskSources+=Join-Path $PSScriptRoot ($taskPair[1]+'.cs')}
 $taskTest=Join-Path $ResultDirectory $taskPair[0]
 & $taskCompiler $taskReferences ('/main:'+$taskPair[1]) ('/out:'+$taskTest) ('/win32manifest:'+(Join-Path $PSScriptRoot 'app.manifest')) $taskSources
 if($LASTEXITCODE -ne 0 -or !(Test-Path -LiteralPath $taskTest)){throw ('测试编译失败：'+$taskPair[1])}
 Copy-Item -LiteralPath ($taskApp+'.config') -Destination ($taskTest+'.config')
 & $taskTest $ResultDirectory
 if($LASTEXITCODE -ne 0){throw ('测试失败：'+$taskPair[1])}
}
