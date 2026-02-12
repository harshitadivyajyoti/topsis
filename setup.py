from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="topsis_harshita_102317208",
    version="0.3",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "numpy",
        "openpyxl"
    ],
    entry_points={
        "console_scripts": [
            "topsis=topsis.topsis:main",
        ],
    },
    author="Harshita",
    description="A Python implementation of Topsis method",
    long_description=long_description,
    long_description_content_type="text/markdown",
    python_requires=">=3.6",
)