from setuptools import setup

setup(
    name = "advancePythonScripts",
    version="0.1",
    description="Python data analysis scripts",
    author="Jacob Sterling",
    author_email="jacob.terling@advance.online",
    packages=["utils", "reports", "remitReaders", "rebates", "MCR", "ACR", "CRMPythonSDK"],
    install_requires=["pywin32", "pandas", "numpy"]
)