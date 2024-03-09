# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

import os
import re
import sys

sys.path.insert(0, os.path.abspath('..'))

version_re = re.compile("__version__ = '(?P<version>[0-9\\.]*)'")
with open('../extract_msg/__init__.py', 'r') as stream:
    contents = stream.read()
match = version_re.search(contents)
__version__ = match.groupdict()['version']


__author__ = 'Destiny Peterson & Matthew Walker'
__year__ = '2024'


project = 'extract-msg Documentation'
copyright = f'{__year__}, {__author__}'
author = __author__
release = __version__

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = ['sphinx.ext.todo', 'sphinx.ext.viewcode', 'sphinx.ext.autodoc']

templates_path = ['_templates']
exclude_patterns = ['_build', 'Thumbs.db', '.DS_Store']

# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = 'sphinx_rtd_theme'
html_static_path = ['_static']
html_style = 'css/override.css'

# -- Options for autodoc -----------------------------------------------------

autoclass_content = 'both'
autodoc_member_order = 'bysource'
