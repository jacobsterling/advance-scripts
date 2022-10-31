from setuptools import setup

setup(
    name = "advancePythonScripts",
    version="0.62",
    description="Python data analysis scripts",
    author="Jacob Sterling",
    author_email="jacob.sterling@advance.online",
    packages=["utils", "reports", "remitReaders", "rebates", "MCR", "ACR", "zohoSDK"],
    install_requires=["pywin32", "pandas", "numpy", "pdfplumber", "xlsxwriter", "openpyxl"]
)