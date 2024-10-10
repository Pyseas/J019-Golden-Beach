
from pathlib import Path
from golden_beach.piping_loads import write_piping_loads


def test_write_loads():

    xlname = Path(__file__).parent.absolute() / 'connector_loads.xlsx'
    outname = Path(__file__).parent.absolute() / 'loadcn.txt'

    write_piping_loads(xlname, outname)


def main():
    test_write_loads()


if __name__ == "__main__":
    main()
