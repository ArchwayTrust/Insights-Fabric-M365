import os
import re

from setuptools import setup, find_packages

HERE = os.path.dirname(os.path.abspath(__file__))


def read_version():
    # Read __version__ from the package without importing it, as importing needs the package's dependencies
    with open(os.path.join(HERE, 'alt_fabric_m365', '__init__.py'), encoding='utf-8') as f:
        match = re.search(r'^__version__\s*=\s*[\'"]([^\'"]+)[\'"]', f.read(), re.MULTILINE)
    if not match:
        raise RuntimeError("Unable to find __version__ in alt_fabric_m365/__init__.py.")
    return match.group(1)


setup(
    name='alt-fabric-m365',
    version=read_version(),
    packages=find_packages(include=["alt_fabric_m365", "alt_fabric_m365.*"]),
    python_requires=">=3.10",
    install_requires=[
        "azure-core",
        "pandas",
        "requests"
    ],
    extras_require={
        # pip install -e .[test], then run pytest
        "test": ["pytest", "pytest-cov", "pyflakes", "setuptools"]
    },
    author="Ben Dobbs",
    author_email="bdobbs@archwaytrust.co.uk",
    description="Collection of useful python code for automation in Microsoft Fabric.",
    long_description=open(os.path.join(HERE, 'README.md'), encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    url="https://github.com/ArchwayTrust/Insights-Fabric-PythonHelper"
)
