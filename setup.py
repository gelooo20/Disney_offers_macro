from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': ['os','xlwings','datetime'], 'excludes': []}

base = 'console'

executables = [
    Executable('main.py', base=base)
]

setup(name='Walt One',
      version = '1.2',
      description = 'A tool for Disney Offers QA used by Wideout/Media Ocean West Coast team',
      options = {'build_exe': build_options},
      executables = executables)

# python setup.py build