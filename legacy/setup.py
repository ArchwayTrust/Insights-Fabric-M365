import os

from setuptools import setup

HERE = os.path.dirname(os.path.abspath(__file__))

# The final release of the old name. It only forwards to alt-fabric-m365; see README.md.
setup(
    name='fabric_python_helper',
    version='0.6.0',
    packages=['fabric_python_helper'],
    python_requires=">=3.10",
    install_requires=["alt-fabric-m365==0.6.*"],
    author="Ben Dobbs",
    author_email="bdobbs@archwaytrust.co.uk",
    description="Renamed alt-fabric-m365. This release only forwards to it.",
    long_description=open(os.path.join(HERE, 'README.md'), encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    url="https://github.com/ArchwayTrust/Insights-Fabric-M365",
    license="BSD-3-Clause",
)
