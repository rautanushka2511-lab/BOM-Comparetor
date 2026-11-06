# PyInstaller hook for numpy to ensure all modules are included
from PyInstaller.utils.hooks import collect_submodules, get_module_file_attribute

hiddenimports = collect_submodules('numpy')
datas = []

# Explicitly add numpy binary modules
try:
    import numpy
    numpy_path = get_module_file_attribute('numpy')
    if numpy_path:
        datas.append((numpy_path.replace('__init__.py', ''), 'numpy'))
except:
    pass

