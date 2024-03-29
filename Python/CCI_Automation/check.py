import subprocess
import sys
import os
from importlib.metadata import version as get_version

requirements_path = os.path.join(os.path.dirname(__file__), 'Setup', 'requirements.txt')
script_to_run = 'CCI_Automation.py'

def check_and_install_requirements(requirements_path):
    with open(requirements_path, 'r') as req_file:
        requirements = [line.strip() for line in req_file if line.strip()]

    packages_to_install = []
    for req in requirements:
        package_name, required_version = req.split('>=')[0], req.split('>=')[1]
        try:
            installed_version = get_version(package_name)
            if installed_version < required_version:
                print(f"{package_name} requires an update.")
                packages_to_install.append(req)
            else:
                print(f"{package_name} already installed with a satisfactory version.")
        except ImportError:
            # Add to list of packages to install
            packages_to_install.append(req)

    if packages_to_install:
        print("Installing or updating packages:", packages_to_install)
        subprocess.check_call([sys.executable, '-m', 'pip', 'install'] + packages_to_install)
    else:
        print("All required packages are installed and up-to-date.")

def run_script(script_name):
    subprocess.check_call([sys.executable, script_name])
    
if __name__ == '__main__':
    check_and_install_requirements(requirements_path)
    run_script(script_to_run)