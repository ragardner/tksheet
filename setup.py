from distutils.core import setup
from setuptools import find_packages

setup(
  name = 'tksheet',
  packages = ['tksheet'] + find_packages(where = 'src'),
  version = '3.5',
  license='MIT',
  description = 'Tkinter table / sheet widget',
  author = 'ragardner',
  author_email = 'ragardner@protonmail.com',
  url = 'https://github.com/ragardner/tksheet',
  download_url = 'https://github.com/ragardner/tksheet/archive/3.5.tar.gz',
  keywords = ['tkinter', 'table', 'widget'],
  install_requires=[],
  classifiers=[
    'Development Status :: 5 - Production/Stable',
    'Intended Audience :: Developers',
    'Topic :: Software Development :: Build Tools',
    'License :: OSI Approved :: MIT License',
    'Programming Language :: Python :: 3.7',
  ],
)
