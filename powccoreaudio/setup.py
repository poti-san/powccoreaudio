from setuptools import find_packages, setup

setup(
    name="powccoreaudio",
    version="0.0.1",
    install_requires=("comtypes", "powc", "powcpropsys", "powcdeviceprop"),
    packages=find_packages(where="src"),
    package_dir={"": "src"},
)
