from distutils.core import setup
setup(
  name = 'tksheet',         # How you named your package folder (MyLib)
  packages = ['tksheet'],   # Chose the same as "name"
  version = '1.1',      # Start with a small number and increase it with every change you make
  license='MIT',        # Chose a license from here: https://help.github.com/articles/licensing-a-repository
  description = 'Tkinter table / sheet widget',   # Give a short description about your library
  author = 'citizen2077',                   # Type in your name
  author_email = 'citizen2077@gmail.com',      # Type in your E-Mail
  url = 'https://github.com/citizen2077/tksheet',   # Provide either the link to your github or to your website
  download_url = 'https://github.com/citizen2077/tksheet/archive/1.1.tar.gz',    # I explain this later on
  keywords = ['tkinter', 'table', 'widget'],   # Keywords that define your package best
  install_requires=[],
  classifiers=[
    'Development Status :: 5 - Production/Stable',      # Chose either "3 - Alpha", "4 - Beta" or "5 - Production/Stable" as the current state of your package
    'Intended Audience :: Developers',      # Define that your audience are developers
    'Topic :: Software Development :: Build Tools',
    'License :: OSI Approved :: MIT License',   # Again, pick a license
    'Programming Language :: Python :: 3.7',
  ],
)
