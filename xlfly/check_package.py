import pkg_resources
import subprocess
import sys


def check_requirements(rqm_txt):
    if rqm_txt is None:
        return

    requirements = rqm_txt.split(" ")

    requirements = [req.strip() for req in requirements]
    installed_packages = {pkg.key: pkg.version for pkg in pkg_resources.working_set}

    missing_packages = []
    for req in requirements:
        package_name = (
            req.split("==")[0].split(">=")[0].split("<=")[0].split("~=")[0].strip()
        )
        if package_name in installed_packages:
            try:
                pkg_resources.require(req)
            except pkg_resources.VersionConflict as e:
                print(
                    f"{package_name} is installed but does not meet the version requirement: {req}"
                )
                print(f"Installed version: {installed_packages[package_name]}")
                missing_packages.append(req)
        else:
            print(f"{package_name} is not installed")
            missing_packages.append(req)

    return missing_packages


def install_packages(packages):
    if packages:
        print("\nInstalling missing packages...")
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + packages)
    else:
        print("\nAll packages are already installed and meet the version requirements.")


def main():
    requirements_file = "requirements.txt"
    missing_packages = check_requirements(requirements_file)
    install_packages(missing_packages)


if __name__ == "__main__":
    main()
