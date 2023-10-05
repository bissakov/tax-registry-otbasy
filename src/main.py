import sys
from os.path import abspath, dirname

sys.path.append(dirname(dirname(abspath(__file__))))


if __name__ == '__main__':
    from src.robot import run
    run()
