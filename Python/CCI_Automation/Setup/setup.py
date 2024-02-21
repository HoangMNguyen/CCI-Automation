import subprocess
import sys
import ensurepip
from packaging import version

def install_dependency(package):
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])

def check_and_install_dependencies():
    # Ensure pip is installed
    ensurepip.bootstrap()

    # Check and install dependencies from requirements.txt
    try:
        with open('requirements.txt', 'r') as file:
            dependencies = file.read().splitlines()
            
            installed_packages = subprocess.check_output([sys.executable, '-m', 'pip', 'freeze'])
            installed_packages = installed_packages.decode('utf-8').splitlines()
            
            for dependency in dependencies:
                package_name, required_version = dependency.split('==')
                installed_version = next((package.split('==')[1] for package in installed_packages if package.startswith(package_name + '==')), None)
                if installed_version is None or version.parse(installed_version) < version.parse(required_version):
                    print(f'Installing {package_name}...')
                    install_dependency(dependency)
                    print(f'{package_name} installed successfully.')
                else:
                    print(f'{package_name} is already installed.')
                    
    except FileNotFoundError:
        print('requirements.txt not found.')

if __name__ == '__main__':
    check_and_install_dependencies()
