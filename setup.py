from setuptools import setup

setup(
    name = "advancePythonScripts",
    version="0.2",
    description="Python data analysis scripts",
    author="Jacob Sterling",
    author_email="jacob.sterling@advance.online",
    packages=["utils", "reports", "remitReaders", "rebates", "MCR", "ACR", "CRMPythonSDK"],
    install_requires=["pywin32", "pandas", "numpy"]
)