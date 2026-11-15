$ErrorActionPreference = 'Stop'
$executable = Join-Path $PSScriptRoot '../Website/bin/Release/net481/Website.exe'
$stdout = Join-Path $env:RUNNER_TEMP 'website-launcher-stdout.txt'
$stderr = Join-Path $env:RUNNER_TEMP 'website-launcher-stderr.txt'
$launcher = $null
$hugoPid = $null
try {
    $launcher = Start-Process -FilePath $executable -ArgumentList '--no-browser' -PassThru `
        -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    $deadline = (Get-Date).AddSeconds(60)
    $address = $null
    while ((Get-Date) -lt $deadline) {
        $launcher.Refresh()
        if ($launcher.HasExited) { throw "Website launcher exited with code $($launcher.ExitCode)." }
        if (Test-Path $stdout) {
            $output = Get-Content $stdout -Raw
            if ($output -match 'Web Server is available at (http://localhost:\d+/)') {
                $address = $Matches[1]
                break
            }
        }
        Start-Sleep -Milliseconds 250
    }
    if (-not $address) { throw 'Website launcher did not report a ready Hugo server.' }
    $response = Invoke-WebRequest -Uri ($address + 'en/') -UseBasicParsing -TimeoutSec 10
    if ($response.StatusCode -ne 200) { throw "Hugo returned HTTP $($response.StatusCode)." }
    $children = @(Get-CimInstance Win32_Process -Filter "ParentProcessId = $($launcher.Id)" |
        Where-Object { $_.Name -ieq 'hugo.exe' })
    if ($children.Count -ne 1) { throw "Expected one Hugo child process; found $($children.Count)." }
    $hugoPid = $children[0].ProcessId

    $launcher.Kill()
    $launcher.WaitForExit()
    $deadline = (Get-Date).AddSeconds(10)
    while ((Get-Date) -lt $deadline -and (Get-Process -Id $hugoPid -ErrorAction SilentlyContinue)) {
        Start-Sleep -Milliseconds 100
    }
    if (Get-Process -Id $hugoPid -ErrorAction SilentlyContinue) {
        throw "Hugo process $hugoPid remained running after the launcher was stopped."
    }
    Write-Host "Hugo served $address and exited with its launcher."
}
catch {
    if (Test-Path $stdout) { Get-Content $stdout | Write-Host }
    if (Test-Path $stderr) { Get-Content $stderr | Write-Host }
    throw
}
finally {
    if ($launcher) {
        $launcher.Refresh()
        if (-not $launcher.HasExited) {
            $launcher.Kill()
            $launcher.WaitForExit()
        }
        $launcher.Dispose()
    }
    if ($hugoPid -and (Get-Process -Id $hugoPid -ErrorAction SilentlyContinue)) {
        Stop-Process -Id $hugoPid
    }
}
