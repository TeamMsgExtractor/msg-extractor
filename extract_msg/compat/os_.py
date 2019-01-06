from os import *
import sys

if sys.version_info[0] >= 3:
	if not hasattr(os, 'getcwdu'):
		os.getcwdu = os.getcwd
