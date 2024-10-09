import subprocess
import pandas as pd
import os
from datetime import datetime
import pkg_resources

def run_command(command):
    try:
        result = subprocess.run(command, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        return result.stdout.strip().splitlines()
    except subprocess.CalledProcessError as e:
        print(f"Error running command '{command}': {e}")
        return []

rpm_command = "rpm -qa --queryformat '%{NAME} %{VERSION} %{INSTALLTIME:date}\n'"
rpm_packages = run_command(rpm_command)

pip_command = "pip freeze"
python_packages = run_command(pip_command)

def get_python_package_install_date(package_name):
    try:
        dist = pkg_resources.get_distribution(package_name)
        install_location = dist.location
        dist_info = dist.egg_info or dist.get_metadata('RECORD')
        if dist_info:
            full_path = os.path.join(install_location, dist_info)
            install_time = os.path.getmtime(full_path)
            install_date = datetime.fromtimestamp(install_time).strftime('%Y-%m-%d')
            return install_date
    except Exception as e:
        return "Unknown"

df_rpm_packages = pd.DataFrame([pkg.split(maxsplit=2) for pkg in rpm_packages], columns=["Package", "Version", "Install Date"])

python_packages_with_dates = []
skipped_packages = []

for pkg in python_packages:
    try:
        if "==" in pkg:
            package_name, version = pkg.split("==")
            install_date = get_python_package_install_date(package_name)
            python_packages_with_dates.append([package_name, version, install_date])
        else:
            skipped_packages.append(pkg)
            print(f"Skipping invalid package format: {pkg}")
    except ValueError:
        skipped_packages.append(pkg)
        print(f"Skipping malformed package entry: {pkg}")

df_python_packages = pd.DataFrame(python_packages_with_dates, columns=["Python Package", "Version", "Install Date"])

df_skipped_packages = pd.DataFrame(skipped_packages, columns=["Skipped Packages"])

with pd.ExcelWriter('server_packages_with_dates_and_env.xlsx', engine='openpyxl') as writer:
    df_rpm_packages.to_excel(writer, sheet_name='RPM Packages', index=False)
    df_python_packages.to_excel(writer, sheet_name='Python Packages', index=False)
    if not df_skipped_packages.empty:
        df_skipped_packages.to_excel(writer, sheet_name='Skipped Packages', index=False)

print("Excel file 'server_packages_with_dates_and_env.xlsx' created successfully.")