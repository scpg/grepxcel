"""Check what XML parser openpyxl 3.x uses internally."""
import sys
sys.path.insert(0, '.')

import inspect
import openpyxl
import openpyxl.reader.excel as reader

src = inspect.getsource(reader)

print(f'openpyxl version : {openpyxl.__version__}')
print(f'uses defusedxml  : {"defusedxml" in src}')
print(f'uses ElementTree : {"ElementTree" in src or "etree" in src.lower()}')

# Also check the XML reader module
import openpyxl.xml.functions as xmlfuncs
src2 = inspect.getsource(xmlfuncs)
print(f'\nxml/functions.py:')
print(f'  defusedxml     : {"defusedxml" in src2}')
print(f'  ElementTree    : {"ElementTree" in src2}')

# Show the actual imports used
for line in src2.splitlines():
    if 'import' in line.lower() and ('xml' in line.lower() or 'etree' in line.lower()):
        print(f'  {line.strip()}')

try:
    import defusedxml
    print(f'\ndefusedxml installed: {defusedxml.__version__}')
except ImportError:
    print('\ndefusedxml NOT installed')
