from setuptools import setup, find_packages

setup(
    name = "advanceScripts",
    version="0.1",
    description="Python data analysis scripts",
    author="Jacob Sterling",
    author_email="jacob.sterling@advance.online",
    packages=["utils", "remitReaders", "reports"],
    install_requires=["pywin32", "pandas", "numpy", "pdfplumber", "xlsxwriter", "openpyxl", "tabula"]
)