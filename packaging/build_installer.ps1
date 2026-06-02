param(
    [string]$Version = ""
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $scriptDir\..\

try {
    $exeName = "SafeSalesSMS.exe"
    $exePath = Join-Path "dist" $exeName
    $iconPath = Join-Path "assets" "SafeSalesSMS.ico"

    if (-not (Test-Path $iconPath)) {
        throw "App icon not found: $iconPath"
    }

    Write-Host "Building PyInstaller executable..."
    python -m PyInstaller `
        --noconfirm `
        --clean `
        --windowed `
        --onefile `
        --icon $iconPath `
        --name "SafeSalesSMS" `
        --version-file packaging/windows_version_info.txt `
        --collect-submodules win32more `
        --collect-binaries win32more `
        --collect-data win32more `
        --collect-submodules easysms `
        --collect-submodules excel `
        --collect-submodules ui `
        ui\__main__.py

    $iscc = Get-Command iscc -ErrorAction SilentlyContinue
    if (-not $iscc) {
        $defaultIsccPath = Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"
        if (Test-Path $defaultIsccPath) {
            $iscc = Get-Item $defaultIsccPath
        }
    }
    if (-not $iscc) {
        throw "Inno Setup compiler not found. Install it by running `choco install innosetup -y`."
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
    & $iscc "/DMyAppVersion=$version" "packaging\installer.iss"
}
finally {
    Pop-Location
}
