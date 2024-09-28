from setuptools import setup, find_packages

VERSION = '0.10.2'
DESCRIPTION = 'Pandas Code Importer with GUI Interface simplifies the process of adding Pandas code into Python scripts. Through a user-friendly GUI, users can easily select and insert Pandas functions, automating repetitive tasks and enhancing productivity in data analysis projects.'
LONG_DESCRIPTION = '''
PyImport is an innovative tool designed to enhance the coding experience by simplifying the integration of the Pandas library into Python scripts. This powerful program, equipped with an intuitive Graphical User Interface (GUI), enables users to efficiently insert Pandas-related functions and code snippets into their projects without the need for manual coding. By automating this process, PyImport streamlines workflows, reduces human error, and accelerates data manipulation tasks for developers and data analysts alike.

The core feature of PyImport is its user-friendly GUI, which allows both beginners and experienced Python users to easily select from a comprehensive list of Pandas modules. Whether you’re working on basic data cleaning or advanced statistical analyses, PyImport lets you quickly browse and choose the relevant Pandas functions, preview the code, and inject it directly into your Python script. This significantly reduces the time spent on writing repetitive code and ensures accuracy by eliminating potential syntax errors.

In addition to importing Pandas modules, PyImport offers customization options. Users can tailor the imported code to fit their project needs, from adjusting function parameters to modifying data structures. The program is built to enhance flexibility while maintaining simplicity, making it the perfect tool for projects of any scale—whether you're just starting out with Python or managing large datasets.

PyImport not only supports rapid prototyping and testing but also encourages better coding practices by automating repetitive tasks and maintaining consistency in code quality. Its GUI interface is lightweight and easy to navigate, allowing users to focus more on the logic and analysis of their data rather than spending time typing and debugging Pandas imports.

For developers who work extensively with data, PyImport provides a bridge between productivity and precision. It’s the ultimate assistant for data manipulation, empowering users to harness the full capabilities of the Pandas library without the overhead of manual imports.'''

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
    install_requires=['pandas', 'openpyxl', 'pillow'],
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
