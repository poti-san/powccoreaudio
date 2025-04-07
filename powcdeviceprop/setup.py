from setuptools import find_packages, setup

setup(
    name="powcdeviceprop",
    version="0.0.1",
    install_requires=("comtypes", "powc", "powcpropsys"),
    packages=find_packages(where="src"),
    package_dir={"": "src"},
)
