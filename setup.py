import glob
import os
from setuptools import setup

import ExtractMsg

# read in the description from README
with open("README.md") as stream:
    long_description = stream.read()

github_url = 'https://github.com/mattgwwalker/msg-extractor'

# read in the dependencies from the virtualenv requirements file
dependencies = []
filename = os.path.join("REQUIREMENTS")
with open(filename, 'r') as stream:
    for line in stream:
        package = line.strip().split('#')[0]
        if package:
            dependencies.append(package)

setup(
    name="ExtractMsg",
    version=ExtractMsg.__version__,
    description="Extracts emails and attachments saved in Microsoft Outlook's .msg files",
    long_description=long_description,
    url=github_url,
    download_url="%s/archives/master" % github_url,
    author='Matthew Walker',
    author_email='mattgwwalker at gmail.com',
    license='GPL',
    scripts=["ExtractMsg.py"],
    py_modules=["ExtractMsg"],
    install_requires=dependencies,
)
