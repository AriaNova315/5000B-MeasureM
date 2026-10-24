# Run on an ONLINE Windows x64 machine from the delivered source tree.
$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
if ($env:OS -ne 'Windows_NT') { throw 'Run this script on Windows.' }
if ((Test-Path '.prepare-venv') -or (Test-Path 'wheelhouse')) {
    throw 'Use a fresh source directory without .prepare-venv or wheelhouse.'
}
py -3.14 -c "import sys,struct,sysconfig; assert sys.version_info[:2] == (3,14) and struct.calcsize('P') == 8 and not sysconfig.get_config_var('Py_GIL_DISABLED')"
if ($LASTEXITCODE -ne 0) {
    throw 'Install standard CPython 3.14 x64 first.'
}
py -3.14 -m venv .prepare-venv
if ($LASTEXITCODE -ne 0) { throw 'Could not create preparation environment.' }
& '.\.prepare-venv\Scripts\python.exe' -m pip --disable-pip-version-check --no-cache-dir download --only-binary=:all: --dest wheelhouse -r requirements-windows.lock.txt
if ($LASTEXITCODE -ne 0) { throw 'Wheel download failed. Do not deliver an incomplete wheelhouse.' }
Get-ChildItem 'wheelhouse\*.whl' | ForEach-Object {
    $hash = Get-FileHash $_.FullName -Algorithm SHA256
    '{0}  {1}' -f $hash.Hash.ToLowerInvariant(), $_.Name
} | Set-Content 'wheelhouse\SHA256SUMS.txt' -Encoding Ascii
& '.\.prepare-venv\Scripts\python.exe' -c "import sys,platform; print(sys.version); print(platform.platform())" | Set-Content 'wheelhouse\PREPARATION.txt' -Encoding Ascii
Write-Host 'Wheelhouse prepared. Follow WINDOWS_DEPLOYMENT.md for a clean offline installation test.'
