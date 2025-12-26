# pptx-template

[![Test](https://github.com/m3dev/pptx-template/actions/workflows/test.yml/badge.svg)](https://github.com/m3dev/pptx-template/actions/workflows/test.yml)
[![Python](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-Apache%202.0-green)](LICENSE)

## Overview

pptx-template is a PowerPoint presentation builder.

This helps your routine reporting work that have many manual copy-paste from excel chart to powerpoint, or so.

  - Building a new powerpoint presentation file from a "template" pptx file which contains "id"
  - Import some strings and CSV data which is defined in a JSON config file or a Python dict
  - "id" in pptx template is expressed as a tiny DSL, like "{sales.0.june.us}"
  - Requires Python 3.10+, pandas, python-pptx
  - For now, only UTF-8 encoding is supported for json, csv

### Text substitution

<img src="docs/01.png?raw=true" width="80%" />

### CSV Import

<img src="docs/02.png?raw=true" width="80%" />

## Getting started

```bash
pip install pptx-template
echo '{ "slides": [ { "greeting" : "Hello!!" } ] }' > model.json

# prepare your template file (test.pptx) which contains "{greeting}" in somewhere

pptx_template --out out.pptx --template test.pptx --model model.json
```

## Development

### Requirements

- Python 3.10, 3.11, 3.12, or 3.13
- [uv](https://docs.astral.sh/uv/) (recommended) or pip

### Installation

```bash
git clone https://github.com/m3dev/pptx-template.git
cd pptx-template

# Using uv (recommended)
uv sync --extra dev

# Or using pip
pip install -e ".[dev]"
```

### Run Tests

```bash
# Using uv
uv run --extra dev pytest

# Or with a specific Python version
uv run --python 3.13 --extra dev pytest

# Using pip (after installation)
pytest
```

### Run via CLI

```bash
# Using uv
uv run pptx_template \
  --template test/data3/in.pptx \
  --model test/data3/in.xlsx \
  --out test/data3/out.pptx \
  --debug

# After pip install
pptx_template \
  --template test/data3/in.pptx \
  --model test/data3/in.xlsx \
  --out test/data3/out.pptx \
  --debug
```

### Run with REPL (for development)

```bash
uv run python
```

```python
import sys
from importlib import reload
import pptx_template.cli as cli

# Set arguments
sys.argv = ['dummy.py', '--out', 'test/data3/out.pptx', '--template', 'test/data3/in.pptx', '--model', 'test/data3/in.xlsx', '--debug']

# Run
cli.main()

# After modifying source code, reload and run again
reload(sys.modules.get('pptx_template.xlsx_model'))
reload(sys.modules.get('pptx_template.text'))
reload(sys.modules.get('pptx_template.table'))
reload(sys.modules.get('pptx_template.chart'))
reload(sys.modules.get('pptx_template.core'))
reload(sys.modules.get('pptx_template.cli'))
cli.main()
```

### Deployment Process

1. Create a feature branch
2. Implement the new feature
3. Test with all versions of Python (3.10, 3.11, 3.12, 3.13)
4. Push changes to the feature branch
5. Create a Pull Request on GitHub
6. Request a code review
7. Verify QA
8. Merge Pull Request
9. Upload to PyPI (Only for PyPI repository administrators)

### Upload to PyPI

```bash
# Build
uv build

# Upload to PyPI
uv publish
```
