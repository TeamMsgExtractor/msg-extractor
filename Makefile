PYTHON = python
ARGS =

# In not in a virtualenv, add --user options for install commands.
INSTALL_OPTS = `$(PYTHON) -c "import sys; print('' if hasattr(sys, 'real_prefix') else '--user')"`

install:  ## Install this package as current user in "edit" mode.
	$(PYTHON) setup.py develop $(INSTALL_OPTS)

test:  ## Run tests.
	$(PYTHON) tests.py

upload:  ## Upload source tarball on PYPI. Requires a .pypirc file in the home dir.
	$(PYTHON) setup.py sdist upload

clean:  ## Remove all build files.
	rm -rf `find . -type d -name __pycache__ \
		-o -type f -name \*.bak \
		-o -type f -name \*.orig \
		-o -type f -name \*.pyc \
		-o -type f -name \*.pyd \
		-o -type f -name \*.pyo \
		-o -type f -name \*.rej \
		-o -type f -name \*.so \
		-o -type f -name \*.~ \
		-o -type f -name \*\$testfn`
	rm -rf \
		*.core \
		*.egg-info \
		*\$testfn* \
		.coverage \
		.tox \
		build/ \
		dist/ \
		docs/_build/ \
		htmlcov/ \
		venv \
		2013-11-18_1026* \
		tmp/

help: ## Display callable targets.
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-20s\033[0m %s\n", $$1, $$2}'
