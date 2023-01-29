import os
from setuptools import setup
import re


# A handful of variables that are used a couple of times.
github_url = 'https://github.com/TeamMsgExtractor/msg-extractor'
main_module = 'extract_msg'

# Read in the description from README.
with open('README.rst', 'rb') as stream:
    long_description = stream.read().decode('utf-8').replace('\r', '')

# Get the version string this way to avoid issues with modules not being
# installed before setup.
version_re = re.compile("__version__ = '(?P<version>[0-9\\.]*)'")
with open('extract_msg/__init__.py', 'r') as stream:
    contents = stream.read()
match = version_re.search(contents)
version = match.groupdict()['version']

# Read in the dependencies from the virtualenv requirements file.
dependencies = []
filename = os.path.join('requirements.txt')
with open(filename, 'r') as stream:
    for line in stream:
        package = line.strip().split('#')[0]
        if package:
            dependencies.append(package)

setup(
    name=main_module,
    version=version,
    description="Extracts emails and attachments saved in Microsoft Outlook's .msg files",
    long_description=long_description,
    long_description_content_type='text/x-rst',
    url=github_url,
    download_url='%s/archives/master' % github_url,
    author='Destiny Peterson & Matthew Walker',
    author_email='arceusthe@gmail.com, mattgwwalker@gmail.com',
    license='GPL',
    packages=[main_module],
    py_modules=[main_module],
    entry_points={
        'console_scripts': ['extract_msg = extract_msg.__main__:main',]
    },
    include_package_data=True,
    install_requires=dependencies,
    python_requires='>=3.8',
)
