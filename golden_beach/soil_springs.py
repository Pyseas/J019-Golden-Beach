
import pandas as pd
from dataclasses import dataclass
from pathlib import Path

PATH = Path('.')


@dataclass
class SoilRanges:
    """Excel ranges for T-Z, Q-Z and P-Y data.

    Attributes:
        tz (str): Excel range (e.g. 'B10:H30') for T-Z data.
        qz (str): Excel range for Q-Z data.
        py (str): Excel range for P-Y data.

    """
    tz: str
    qz: str
    py: str


def read_range(xlpath: Path, start_row: int, nrow: int, cols: str) -> list[pd.DataFrame]:

    df = pd.read_excel(
        xlpath, skiprows=start_row, nrows=nrow-1, usecols=(cols), header=None)
    df = df.round(2)

    ndata = len(df.columns) - 2
    columns = ['Depth', 'Point'] + [i for i in range(1, ndata + 1)]
    df.columns = columns
    df.drop(['Point'], axis=1, inplace=True)

    a = df.iloc[::2, :].reset_index(drop=True).copy()
    b = df.iloc[1::2, :].reset_index(drop=True).copy()
    b.loc[:, 'Depth'] = a.loc[:, 'Depth']
    
    return [a, b]


def get_intro() -> str:

    outstr = 'PSIOPT +ZMN   Y       EX0.002540  0.0001 20               S3 100        7.849047\n'
    outstr += 'PLTRQ SD   DTE  RTE            TSE  DAE  AL   AS   UC             XH\n'
    outstr += 'PLTLC 1001 2001 2002 2003 2004 2005 2006 2007 2008 2009 2010 2011 2012 2013\n'
    outstr += 'PLTLC 2014 2015 2016 3001 3002 3003 3004 3005 3006 3007 3008 3009 3010 3011\n'
    outstr += 'PLTLC 3012 3013 3014 3015 3016\n'
    outstr += '*\n'
    outstr += 'LCSEL IN        1001 2001 2002 2003 2004 2005 2006 2007 2008 2009 2010 2011\n'
    outstr += 'LCSEL IN        2012 2013 2014 2015 2016 3001 3002 3003 3004 3005 3006 3007\n'
    outstr += 'LCSEL IN        3008 3009 3010 3011 3012 3013 3014 3015 3016\n'
    outstr += 'PLGRUP\n'
    outstr += '*PLUGGED\n'
    outstr += '*PLGRUP PL1          91.00  2.5419.9957.997924.821   30.00              1.00.6503\n'
    outstr += '*UNPLUGGED\n'
    outstr += 'PLGRUP PL1        U 91.00  2.5419.9957.997924.821   30.00              1.00.0706\n'
    outstr += 'PILE\n'
    outstr += 'PILE  A031     PL1                           3000.                  SOL1\n'
    outstr += 'PILE  A032     PL1                           3000.                  SOL1\n'
    outstr += 'PILE  A033     PL1                           3000.                  SOL1\n'
    outstr += 'PILE  A034     PL1                           3000.                  SOL1\n'
    outstr += 'SOIL\n'

    return outstr


def get_tz_str(t: pd.DataFrame, z: pd.DataFrame, title: str) -> str:

    t = t.drop_duplicates('Depth', keep='first').reset_index(drop=True)
    z = z.drop_duplicates('Depth', keep='first').reset_index(drop=True)
    depths = t.loc[:, 'Depth'].to_list()

    ndep = len(depths)
    outstr = f'*{title}\n'
    outstr += 'SOIL TZAXIAL HEAD' + f'{ndep: >3}' + ' '*20 + 'SOL1\n'

    for idep in range(ndep):
        t_kpa = t.iloc[idep].to_numpy()[1:]
        z_mm = z.iloc[idep].to_numpy()[1:]
        npt = len(t_kpa)

        outstr += 'SOIL T-Z     SLOCSM  ' + f'{npt: >2} '
        ztop = f'{depths[idep]:6.1f}'
        if idep == ndep - 1:
            zbot = ' '*6
        else:
            zbot = f'{depths[idep + 1]:6.1f}'
        outstr += ztop + zbot
        outstr += '  ' + f'{1.0: >5}\n'

        outstr += 'SOIL         T-Z '
        for t0, z0 in zip(t_kpa, z_mm):
            if 0.0 < t0 < 100:
                t_str = f'{t0/10000:.2e}'.replace('e-0', 'e-').replace('e-', '-')
            else:
                t_str = f'{t0/10000:6.4f}'
            outstr += f'{t_str: >6}'
            outstr += f'{z0/10:6.1f}'
        outstr += '\n'

    return outstr


def get_qz_str(q: pd.DataFrame, z: pd.DataFrame, thk: float) -> str:

    q = q.drop_duplicates('Depth', keep='first').reset_index(drop=True)
    z = z.drop_duplicates('Depth', keep='first').reset_index(drop=True)
    depths = q.loc[:, 'Depth'].to_list()

    ndep = len(depths)
    outstr = 'SOIL BEARING HEAD' + f'{ndep: >3}' + ' '*20 + 'SOL1\n'

    for idep in range(ndep):
        q_kpa = q.iloc[idep].to_numpy()[1:]
        z_mm = z.iloc[idep].to_numpy()[1:] * thk
        npt = len(q_kpa)

        outstr += 'SOIL BEAR    SLOCSM  ' + f'{npt: >2} '
        ztop = f'{depths[idep]:6.1f}'
        if idep == ndep - 1:
            zbot = f'{depths[idep] + 1.0:6.1f}'
        else:
            zbot = f'{depths[idep + 1]:6.1f}'
        outstr += ztop + zbot
        outstr += '  ' + f'{1.0: >5}'

        ipt = 0
        for q0, z0 in zip(q_kpa, z_mm):
            if ipt % 5 == 0:
                outstr += '\nSOIL         T-Z '
            if 0.0 < q0 < 100:
                q_str = f'{q0/10000:.2e}'.replace('e-0', 'e-').replace('e-', '-')
            else:
                q_str = f'{q0/10000:6.4f}'
            outstr += f'{q_str: >6}'
            outstr += f'{z0:6.1f}'
            ipt += 1
        outstr += '\n'

    return outstr


def get_py_str(p: pd.DataFrame, y: pd.DataFrame) -> str:

    depths = p.loc[:, 'Depth'].to_list()

    ndep = len(depths)
    outstr = 'SOIL LATERAL HEAD' + f'{ndep: >3}'
    outstr += '   YEXP   91.       SOL1                NN  10N\n'

    for idep in range(ndep):
        p_kn = p.iloc[idep].to_numpy()[1:]
        y_mm = y.iloc[idep].to_numpy()[1:]
        npt = len(p_kn)

        outstr += 'SOIL P-Y     SLOCSM  ' + f'{npt: >2} '
        ztop = f'{depths[idep]:6.1f}'
        if idep == ndep - 1:
            zbot = ' '*6
        else:
            if depths[idep] == depths[idep + 1]:
                zbot = f'{depths[idep + 1] + 0.001:6.3f}'
            else:
                zbot = f'{depths[idep + 1]:6.1f}'
        outstr += ztop + zbot
        outstr += f'{1.0: >4}'

        ipt = 0
        for p0, y0 in zip(p_kn, y_mm):
            if ipt % 5 == 0:
                outstr += '\nSOIL         P-Y '
            if 0.0 < p0 < 100:
                p_str = f'{p0/100:.2e}'.replace('e-0', 'e-').replace('e-', '-')
            else:
                p_str = f'{p0/100:6.3f}'
            outstr += p_str
            outstr += f'{y0/10:6.3f}'
            ipt += 1
        outstr += '\n'

    return outstr


def split_cell_address(cell_address: str) -> tuple[int, str]:

    for ichar, letter in enumerate(cell_address):
        if letter.isnumeric():
            break

    col = cell_address[:ichar]
    row = int(cell_address[ichar:])

    return row, col


def range_to_ind(rng: str) -> tuple[int, int, str]:
    # Converts Excel range (e.g. 'A2:BA61') to start_row,
    # nrows and cols (e.g. 'A:BA')
    start, end = rng.split(':')

    row0, col0 = split_cell_address(start)
    row1, col1 = split_cell_address(end)

    start_row = row0
    nrows = row1 - row0 + 1
    cols = f'{col0}:{col1}'

    return start_row, nrows, cols


def write_soil_springs(xlname: str, outname: str, ranges: SoilRanges,
                       tz_title: str, B: float) -> None:
    """Write soil springs data from spreadsheet to SACS format file.

    Args:
        xlname (str): Spreadsheet filename.
        outname (str): Output filename.
        ranges (SoilRanges): SoilRanges object.
        tz_title (str): Title to be added as comment at start of T-Z section.
        B (float): For plugged pile = OD, unplugged=WT, in cm

    """

    xlpath = PATH.joinpath(xlname)
    outpath = PATH.joinpath(outname)

    # T-Z
    start_row, nrow, cols = range_to_ind(ranges.tz)
    t, z = read_range(xlpath, start_row, nrow, cols)
    tz_str = get_tz_str(t, z, tz_title)

    # Q-Z
    start_row, nrow, cols = range_to_ind(ranges.qz)
    q, z = read_range(xlpath, start_row, nrow, cols)
    qz_str = get_qz_str(q, z, thk=B)

    # P-Y
    start_row, nrow, cols = range_to_ind(ranges.py)
    p, y = read_range(xlpath, start_row, nrow, cols)
    py_str = get_py_str(p, y)

    outstr = get_intro()
    outstr += tz_str
    outstr += qz_str
    outstr += 'SOIL TORSION HEAD                  5000.SOL1                N\n'
    outstr += py_str
    outstr += 'END\n'

    with open(outpath, 'w') as f:
        f.write(outstr)


def main():

    xlname = 'GB_PLEM_SoilSprings_23Aug2024.xlsx'


if __name__ == "__main__":
    main()
