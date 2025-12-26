pptx-template
=============

.. image:: https://github.com/m3dev/pptx-template/actions/workflows/test.yml/badge.svg
    :target: https://github.com/m3dev/pptx-template/actions/workflows/test.yml

pptx-template is a PowerPoint presentation builder.

This helps your routine reporting work that have many manual copy-paste from excel chart to powerpoint, or so.

  - Building a new powerpoint presentation file from a "template" pptx file which contains "id"
  - Import some strings and CSV data which is defined in a JSON config file or a Python dict
  - "id" in pptx template is expressed as a tiny DSL, like "{sales.0.june.us}"
  - Requires Python 3.10+, pandas, python-pptx
  - For now, only UTF-8 encoding is supported for json, csv

For further information, please visit GitHub: https://github.com/m3dev/pptx-template

Changelog
=========

1.0.0 (2024-12)
---------------

  - Dropped Python 2.x and Python 3.9 and earlier support
  - Added Python 3.10, 3.11, 3.12, 3.13 support
  - Updated dependencies: python-pptx>=1.0.0, pandas>=2.0.0, openpyxl>=3.1.0
  - Migrated from setup.py to pyproject.toml
  - Migrated from pip to uv
  - Migrated CI from Travis CI to GitHub Actions

0.2.9 (2019)
------------

  - Last version supporting Python 2.7 and Python 3.5-3.7

2017.07.18
----------

  - Added "xlsx-mode"
  - Fixed many small bugs
