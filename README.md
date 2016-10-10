ODocxTpl
========

[![License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat)](https://github.com/Ventto/odocxtpl/blob/master/LICENSE)
[![Status](https://img.shields.io/badge/status-experimental-orange.svg?style=flat)](https://github.com/Ventto/odocxtpl/)

*ODocxTpl is a simple tool to generate a word document from .docx template and
YAML variables (using py-docx-tpl).*

**Warning: This tool is experimental.**

Requirements
------------

* *python3.5*
* *python-pip* - The PyPA recommended tool for installing Python packages
* *python-yaml* - Python bindings for YAML, using fast libYAML library

Installation
------------

Install *python-docx-template*:

```
$ pip install docxtpl
```

Before usage:

```
$ chmod +x src/odocxtpl.py
```

Usage
-----

```
$ ./odocxtpl.py file1.docx meta.yaml out.docx
```

* **file1.docx**: Word document containing context {{ vars }}
* **meta.yaml**: Variable values
* **out.docx**: Output file
