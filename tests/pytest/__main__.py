import sys

from . import _runner

if __name__ == "__main__":
    sys.exit(_runner.main(sys.argv[1:]))
