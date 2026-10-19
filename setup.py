from setuptools import setup, find_packages

setup(
    name='fabric_python_helper',
    version="0.3.0",
    packages=find_packages(),
    install_requires=[
        "azure-core",
        "msal",
        "requests"
    ],
    author="Ben Dobbs",
    author_email="bdobbs@archwaytrust.co.uk",
    description="Collection of useful python code for automation in Microsoft Fabric.",
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url="https://github.com/ArchwayTrust/Fabric-PythonHelper"
)