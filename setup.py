from setuptools import setup

setup(
    name="TelephoneTool",
    version="1.0.0",
    description="コールセンター業務効率化ツール",
    author="Your Name",
    packages=["telephonetool"],
    install_requires=[
        "PySide6",
        "gspread",
        "oauth2client",
        "pykakasi"
    ],
    python_requires=">=3.8",
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3.8",
    ],
) 