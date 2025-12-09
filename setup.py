from setuptools import setup, find_packages

setup(
    name="excel_charts",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "pandas>=1.3.0",
        "xlsxwriter>=3.0.0",
    ],
    extras_require={
        "dev": [
            "pytest>=7.0",
            "black>=22.0",
            "flake8>=4.0",
        ],
    },
    python_requires=">=3.9",
    author="Edward T.L.",
    author_email="edward_tl@hotmail.com",
    description="Object-oriented Python library for creating Excel charts using xlsxwriter",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/Edward-TL/excel_charts",
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
    keywords="excel charts xlsxwriter visualization pandas",
)
