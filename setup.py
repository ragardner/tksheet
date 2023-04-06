from setuptools import setup, find_packages
from os import path

this_directory = path.abspath(path.dirname(__file__))
with open(path.join(this_directory, 'README.md'), encoding = 'utf-8') as f:
    long_description = f.read()

setup(
  name = 'tksheet',
  packages = ['tksheet'],
  version = '5.6.8',
  python_requires = '>=3.6',
  license = 'MIT',
  description = 'Tkinter table / sheet widget',
  long_description = long_description,
  long_description_content_type = 'text/markdown',
  author = 'ragardner',
  author_email = 'github@ragardner.simplelogin.com',
  url = 'https://github.com/ragardner/tksheet',
  download_url = 'https://github.com/ragardner/tksheet/archive/5.6.8.tar.gz',
  keywords = ['tkinter', 'table', 'widget', 'sheet', 'grid', 'tk'],
  install_requires = [],
  classifiers = [
    'Development Status :: 5 - Production/Stable',
    'Intended Audience :: Developers',
    'Topic :: Software Development :: Build Tools',
    'License :: OSI Approved :: MIT License'
  ],
)
