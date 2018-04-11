from distutils.core import setup

setup(
    name='pack',
    version='1.0',
    packages=['pack', 'pack.reader','pack.store', 'pack.utils', 'pack.labels'],
    license='Creative Commons Attribution-Noncommercial-Share Alike license',
    package_data={'pack.labels': ['view/*', 'view/font/*']}
)
