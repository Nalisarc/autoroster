from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
buildOptions = dict(packages = [], excludes = [])

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('qt.py', base=base, targetName = 'autoroster.exe')
]

setup(name='AutoRoster',
      version = '0.2',
      description = 'A program for creating roster reports.',
      options = dict(build_exe = buildOptions),
      executables = executables)
