import glob
import os
from setuptools import setup
import re


# a handful of variables that are used a couple of times
github_url = 'https://github.com/mattgwwalker/msg-extractor'
main_module = 'ExtractMsg'
main_script = main_module + '.py'

# read in the description from README
with open("README.md") as stream:
    long_description = stream.read()

# get the version from the ExtractMsg script. This can not be directly
# imported because ExtractMsg imports olefile, which isn't installed yet
version_re = re.compile("__version__ = '(?P<version>[0-9\.]*)'")
with open(main_script, 'r') as stream:
    contents = stream.read()
match = version_re.search(contents)
version = match.groupdict()['version']

# read in the dependencies from the virtualenv requirements file
dependencies = []
filename = os.path.join("REQUIREMENTS")
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
    url=github_url,
    download_url="%s/archives/master" % github_url,
    author='Matthew Walker',
    author_email='mattgwwalker at gmail.com',
    license='GPL',
    scripts=[main_script],
    py_modules=[main_module],
    install_requires=dependencies,
)
