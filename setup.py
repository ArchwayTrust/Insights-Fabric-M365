import os
import re

from setuptools import setup, find_packages

HERE = os.path.dirname(os.path.abspath(__file__))


def read_version():
    # Read __version__ from the package without importing it, as importing needs the package's dependencies
    with open(os.path.join(HERE, 'fabric_python_helper', '__init__.py'), encoding='utf-8') as f:
        match = re.search(r'^__version__\s*=\s*[\'"]([^\'"]+)[\'"]', f.read(), re.MULTILINE)
    if not match:
        raise RuntimeError("Unable to find __version__ in fabric_python_helper/__init__.py.")
    return match.group(1)


setup(
    name='fabric_python_helper',
    version=read_version(),
    packages=find_packages(),
    python_requires=">=3.10",
    install_requires=[
        "azure-core",
        "pandas",
        "requests"
    ],
    extras_require={
        # PemCertificateCredential only. msal is already in the Fabric runtime.
        "msal": ["msal"],
        # pip install -e .[test], then run pytest
        "test": ["pytest", "pytest-cov", "pyflakes", "setuptools", "msal"]
    },
    author="Ben Dobbs",
    author_email="bdobbs@archwaytrust.co.uk",
    description="Collection of useful python code for automation in Microsoft Fabric.",
    long_description=open(os.path.join(HERE, 'README.md'), encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    url="https://github.com/ArchwayTrust/Insights-Fabric-PythonHelper"
)
