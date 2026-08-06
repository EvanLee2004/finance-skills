param(
    [int]$Port = 9222,
    [string]$UserDataDir
)

$target = Join-Path $PSScriptRoot "..\..\jdy-cashflow-export\scripts\start_edge.ps1"
$arguments = @("-Port", $Port)
if ($UserDataDir) {
    $arguments += @("-UserDataDir", $UserDataDir)
}

& $target @arguments
exit $LASTEXITCODE
