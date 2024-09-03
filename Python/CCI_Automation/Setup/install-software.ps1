# Check if Chocolatey is installed
if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
    # Install Chocolatey
    Set-ExecutionPolicy Bypass -Scope Process -Force; 
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12; 
    iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
} else {
    Write-Host "Chocolatey is already installed."
}

# List of software packages to install
$packages = @(
    "wsl2",
    "docker-desktop",
    "vcxsrv"
)

# Install each package
foreach ($package in $packages) {
    choco install $package -y
}

# Update all installed packages (optional)
choco upgrade all -y
