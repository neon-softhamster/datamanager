from setuptools import setup, find_packages

setup(
    name = "datamanager",
    version = "0.6.0",
    packages = find_packages(),
    install_requires = [
        'numpy>=1.26.4',
        'pandas>=2.2.1',
        'scipy>=1.15.2',
    ],
    author = "SoftHamster",
    description = "Easy reading and writing tables and more"
)
