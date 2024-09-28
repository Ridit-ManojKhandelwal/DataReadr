from setuptools import setup, find_packages
import os
import codecs

here = os.path.abspath(os.path.dirname(__file__))

with codecs.open(os.path.join(here, "README.md"), encoding="utf-8") as fh:
    long_description = "\n" + fh.read()

VERSION = '0.11.3'
DESCRIPTION = 'DataReadr GUI Interface simplifies the process of adding Pandas code into Python scripts. Through a user-friendly GUI, users can easily select and insert Pandas functions, automating repetitive tasks and enhancing productivity in data analysis projects.'
LONG_DESCRIPTION = 'An Application that simplifies the process of creating dataframes in python with GUI'

# Setting up
setup(
    name="datareadr",
    version=VERSION,
    author="Manoj Khandelwal",
    author_email="tr.manojkhandelwal@gmail.com",
    description=DESCRIPTION,
    long_description_content_type="text/markdown",
    long_description=LONG_DESCRIPTION,
    packages=find_packages(),
    install_requires=['pandas', 'openpyxl', 'pillow', 'cryptography'],
    keywords=['python', 'pandas', 'gui', 'importer', 'pandas importer', 'import', 'pyimport', 'pyimport-0.9.0'],
    classifiers=[
        "Development Status :: 1 - Planning",
        "Intended Audience :: Developers",
        "Programming Language :: Python",
        "Operating System :: Unix",
        "Operating System :: MacOS :: MacOS X",
        "Operating System :: Microsoft :: Windows",
    ]
)
