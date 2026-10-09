#!/usr/bin/env python3
"""Run native consumer tests using a Python with the optional test dependencies.

The launcher itself needs only the standard library. It never installs packages
or runs PowerPoint; VBA compilation and saved host results are separate checks.
"""
from pathlib import Path
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent
DEPENDENCY_PROBE = 'import olefile; from oletools.olevba import VBA_Parser, decompress_stream'


def find_test_python():
    candidates = [ROOT / '.venv-native/bin/python', sys.executable,
                  shutil.which('python3'), shutil.which('python')]
    checked = set()
    for candidate in candidates:
        if not candidate:
            continue
        # Preserve the venv path: resolving its symlink would lose the venv.
        executable = str(Path(candidate).absolute())
        if executable in checked or not Path(executable).is_file():
            continue
        checked.add(executable)
        try:
            probe = subprocess.run([executable, '-c', DEPENDENCY_PROBE],
                                   capture_output=True, timeout=10)
        except (OSError, subprocess.SubprocessError):
            continue
        if probe.returncode == 0:
            return executable
    return None


def main():
    executable = find_test_python()
    if executable is None:
        print('Native format/installation tests need olefile and oletools.\n'
              'No available Python has these optional test dependencies.\n'
              'From the project directory, run:\n'
              '  python3 -m venv .venv-native\n'
              '  .venv-native/bin/python -m pip install -r native/requirements-test.txt\n'
              'Then run npm test again.', file=sys.stderr)
        return 1
    print('Native PPAM format/installation checks; no PowerPoint or VBA execution.\n'
          'Python: ' + executable, flush=True)
    result = subprocess.run([executable, str(ROOT / 'test/test-native-package.py'),
                             *sys.argv[1:]], cwd=ROOT)
    return result.returncode


if __name__ == '__main__':
    raise SystemExit(main())
