from setuptools import setup, find_packages

setup(
    name='metadata_extractor',
    version='0.1',
    packages=find_packages(),
    install_requires=[
        'olefile',
        'pdfminer.six'
    ]
)
