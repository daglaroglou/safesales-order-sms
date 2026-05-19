param(
    [string]$Version = ""
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $scriptDir\..\

try {
    $exeName = "SafeSalesSMS.exe"
    $exePath = Join-Path "dist" $exeName

    if (-not (Test-Path $exePath)) {
        Write-Host "Building PyInstaller executable..."
        python -m PyInstaller `
            --noconfirm `
            --clean `
            --windowed `
            --onefile `
            --name "SafeSalesSMS" `
            --version-file packaging/windows_version_info.txt `
            --collect-submodules win32more `
            --collect-binaries win32more `
            --collect-data win32more `
            --collect-submodules easysms `
            --collect-submodules excel `
            --collect-submodules ui `
            ui\__main__.py
    }

    $makensis = Get-Command makensis -ErrorAction SilentlyContinue
    if (-not $makensis) {
        throw "NSIS compiler not found. Install NSIS by running `choco install nsis -y` or install it manually."
    }

    if (-not $Version) {
        $version = "0.0.0-dev"
        try {
            $tag = git describe --tags --abbrev=0 2>$null
            if ($LASTEXITCODE -eq 0 -and $tag) {
                $version = $tag.TrimStart("v")
            } else {
                $shortSha = git rev-parse --short HEAD
                if ($LASTEXITCODE -eq 0 -and $shortSha) {
                    $version = "0.0.0-dev.$shortSha"
                }
            }
        } catch {
            # Fallback to default version when git is unavailable.
        }
    } else {
        $version = $Version.TrimStart("v")
    }

    Write-Host "Building installer for version $version"
    & $makensis "/DMAyAppVersion=$version" "packaging\installer.iss"
}
finally {
    Pop-Location
}
