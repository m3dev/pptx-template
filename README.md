# pptx-template [![Build Status](https://travis-ci.org/m3dev/pptx-template.svg?branch=master)](https://travis-ci.org/m3dev/pptx-template)

## Overview

pptx-template is a PowerPoint presentation builder.

This helps your routine reporting work that have many manual copy-paste from excel chart to powerpoint, or so.

  - Building a new powerpoint presentation file from a "template" pptx file which contains "id"
  - Import some strings and CSV data which is defined in a JSON config file or a Python dict
  - "id" in pptx template is expressed as a tiny DSL, like "{sales.0.june.us}"
  - requires python envirionment (3), pandas, python-pptx
  - for now, only UTF-8 encoding is supported for json, csv

### Text substitution

<img src="docs/01.png?raw=true" width="80%" />

### CSV Import

<img src="docs/02.png?raw=true" width="80%" />

## Getting started

TBD

```
$ pip install pptx-template
$ echo '{ "slides": [ { "greeting" : "Hello!!" } ] }' > model.json

# prepare your template file (test.pptx) which contains "{greeting}" in somewhere

$ pptx-template --out out.pptx --template test.pptx --model model.json
```

## Development

### Installation

Install using `pyenv`

```
git clone https://github.com/m3dev/pptx-template.git

pyenv install 3.7.1 # Install Python
pyenv shell 3.7.1 # Create Python shell

venv .venv # Create virtual environment for development
source .venv/bin/activate # Use the virtual environment for development

python setup.py develop         # Setup egg-info folder for development & Install dependencies
pip install -r requirements.txt # Install dependencies
```

### Run with REPL - Use this for development

Launch the Python REPL client

```
cd {project folder}
pyenv shell 3.7.1
python
```

Run the following with the Python REPL

```
import sys
from importlib import reload
import pptx_template.cli as cli


## Argument Settings
## sys.argv = ['{filename.py}', '--out', '{file/path/output.pptx}', '--template', '{file/path/template.pptx}', '--model', '{file/path/data.xlsx}', '--debug']
## Following is an example with test files
sys.argv = ['dummy.py', '--out', 'test/data3/out.pptx', '--template', 'test/data3/in.pptx', '--model', 'test/data3/in.xlsx', '--debug']

## Run the program
cli.main()

## Run the following after modifying the source code
reload(sys.modules.get('pptx_template.xlsx_model'))
reload(sys.modules.get('pptx_template.text'))
reload(sys.modules.get('pptx_template.table'))
reload(sys.modules.get('pptx_template.chart'))
reload(sys.modules.get('pptx_template.core'))
reload(sys.modules.get('pptx_template.cli'))
cli.main()
```

### Run via CLI

```
## pptx_template --out {file/path/output.pptx} --template {file/path/template.pptx} --model {file/path/data.xlsx}  --debug
pptx_template --out test/data3/out.pptx --template test/data3/in.pptx --model test/data3/in.xlsx  --debug
```

### Run Tests

```
pytest
```

### Deployment Process

1. Create a feature branch
2. Implement the new feature
3. Test with all versions of Python
4. Push changes to the feature branch
5. Create a Pull Request on Github
6. Request a code review
7. Verify QA（If you want to verify QA、build the source in your local environment）
8. Merge Pull Request
9. Upload to PyPI（Only for PyPI repository administrators）

### Upload to PyPI

1. Install packages needed for uploading to PyPI

```
pip install wheel
pip install twine
```

2. Compile

```
python setup.py bdist_wheel
```

3. Upload to PyPI

```
twine upload dist/*
```
