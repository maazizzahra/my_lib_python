# setup.py
from setuptools import setup, find_packages

setup(
    name="excel_lib",
    version="0.1",
    packages=find_packages(),
    install_requires=["openpyxl"],
)
