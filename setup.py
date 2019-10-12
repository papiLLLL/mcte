from setuptools import setup, find_packages

with open("README.md") as f:
    readme = f.read()

with open("LICENSE.txt") as f:
    license = f.read()

setup(
    name="mcte",
    version="1.0.0",
    description="Copy multipule csv files to excel.",
    long_description=readme,
    author="Yuitoku",
    author_email="exrecord160@gmail.com",
    install_requires=["openpyxl"],
    url="https://github.com/yuitoku/mcte",
    license=license,
    packages=find_packages(exclude=("csv", "template"))
)