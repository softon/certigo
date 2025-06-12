# CLI Guide

```
pip install pyinstaller

pyinstaller --onefile --name certigo --add-data "config.json;data.xlsx;template.png" certigo.py
pyinstaller --onefile --name certigo certigo.py
pyinstaller --onefile --name certigen certigen.py
```