$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    & node tools/build.mjs
    if ($LASTEXITCODE -ne 0) { throw 'Static frontend build failed.' }
    & node --test 'tools/*.test.mjs'
    if ($LASTEXITCODE -ne 0) { throw 'Frontend checks failed.' }
} finally { Pop-Location }
