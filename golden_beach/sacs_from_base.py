
# import numpy as np
import pandas as pd
import os
from pathlib import Path
from typing import Any

PATH = Path('.')


def mem_str(row: Any, line: str) -> str:

    if len(line) < 41:
        line = f'{line: <41}'

    stress = row['STRESS']
    if row['STRESS'] == -123456:
        stress = line[19:21]

    gap = row['GAP']
    if row['GAP'] == -123456:
        gap = line[21]

    fix_a = str(row['FIX_A']).zfill(6)
    if row['FIX_A'] == -123456:
        fix_a = line[22:28]

    fix_b = str(row['FIX_B']).zfill(6)
    if row['FIX_B'] == -123456:
        fix_b = line[28:34]

    angle = f'{row["ANGLE"]:6.1f}'
    if row['ANGLE'] == -123456:
        angle = line[34:41]

    offset = False
    mem_offset = '\nMEMBER OFFSETS' + ' ' * 21
    offset_option = ' '
    for end in ['A', 'B']:
        for dof in ['X', 'Y', 'Z']:
            val = row[f'OFF_{end}{dof}']
            if val == -123456:
                mem_offset += '      '
            else:
                offset_option = '1'
                offset = True
                if val <= -100000:
                    val_str = 'xxxxxx' # Need exponential format
                elif -100000 < val <= -1000:
                    val_str = f'{val:6.0f}'
                elif -1000 < val <= -100:
                    val_str = f'{val:6.1f}'
                elif -100 < val <= 1000:
                    val_str = f'{val:6.2f}'
                elif 1000 < val <= 10000:
                    val_str = f'{val:6.1f}'
                elif 10000 < val <= 1000000:
                    val_str = f'{val:6.0f}'
                else:
                    val_str = 'xxxxxx' # Need exponential format
                mem_offset += val_str

    outstr = 'MEMBER' + offset_option
    outstr += row['ID'] + ' ' + row['GRUP']
    outstr += stress + gap + fix_a + fix_b + angle + line[41:]
    if offset:
        outstr += mem_offset

    return outstr


def jnt_coords(row: dict) -> str:

    outstr = 'JOINT ' + row['JNT'] + ' '
    if row['X'] == -123456:
        # If coordinate cells in spreadsheet are blank
        outstr = 'xxx'
    else:
        for var in ['X', 'Y', 'Z']:
            outstr += f'{row[var]:7.3f}'

    return outstr


def jnt_str(row: dict, line:str) -> str:

    outstr = line.strip()
    fixity = ''
    elasti = ''
    if row['Special'] != -123456:
        fixity = row['Special']
    else:
        for var in ['FX', 'FY', 'FZ', 'FRX', 'FRY', 'FRZ']:
            if row[var] == -123456:
                fixity += '0'
                elasti += ' ' * 7
            elif row[var] == 'F':
                fixity += '1'
                elasti += ' ' * 7
            elif row[var] > 0:
                fixity += '1'
                stiff = float(row[var])
                if stiff >= 100000:
                    elasti += f'{stiff:7.0f}'
                elif 1000 <= stiff < 100000:
                    elasti += f'{stiff:7.1f}'
                else:
                    elasti += f'{stiff:7.2f}'

    if (fixity == 'PILEHD') or ('1' in fixity):
        outstr += ' '*22 + fixity
    if elasti.strip():
        outstr += '\n' + line[:11] + elasti + ' ELASTI'

    return outstr


def lcomb_str(lcomb: pd.DataFrame) -> str:

    loadcns = list(lcomb['LOADCN'])

    outstr = 'LCOMB\n'
    for col in lcomb.columns[3:]:
        lc = lcomb[col].dropna()
        lcname = lc.name
        if isinstance(lcname, int):
            lcname = f'{lcname:04d}'
        outstr += f'LCOMB {lcname: >4} '
        for ind, val in lc.items():
            loadcn = loadcns[ind]
            if isinstance(loadcn, int):
                loadcn = f'{loadcn:04d}'
            outstr += f'{loadcn: >4}{val: >6}'

        outstr += '\n'

    return outstr


def read_loadfiles(loadfiles: list[str]) -> str:

    outstr = ''
    for loadfile in loadfiles:
        loadpath = PATH.joinpath('insert_files', loadfile)
        if os.path.getsize(loadpath) < 10:
            continue
        with open(loadpath, 'r') as f:
            for line in f:
                outstr += line

    return outstr


def make_new_model(xlname: str, basename: str, newname: str):
    """Create new SACS model from a base model and a spreadsheet.

    Add new or modify existing joints/members, add/remove basic load conditions,
    add load combinations.

    Args:
        xlname (str): Spreadsheet filename.
        basename (str): SACS base model filename.
        newname (str): SACS output filename.

    """

    xlpath = PATH.joinpath(xlname)

    titles = pd.read_excel(xlpath, sheet_name='TITLE', header=None)
    title = titles.iloc[0][1]
    analysis_type = titles.iloc[1][1]

    jnts = pd.read_excel(xlpath, sheet_name='joints', skiprows=1, usecols='A:K',
                         converters={'Special': str})
    jnts.fillna(-123456, inplace=True)
    njnt = len(jnts.index)
    new_joints = []
    for ijnt in range(njnt):
        jnt_id = jnts['JNT'].iloc[ijnt]
        if jnt_id == -123456:
            njnt = ijnt
        else:
            new_joints.append(jnt_id)
    # jnts.set_index('JNT', inplace=True)

    mems = pd.read_excel(
        xlpath, sheet_name='members', skiprows=1, usecols='A:N',
        converters={'FIX_A': str, 'FIX_B': str})
    mems.fillna(-123456, inplace=True)
    mems['ID'] = mems['A'] + mems['B']
    nmem = len(mems.index)
    new_mems = []
    for imem in range(nmem):
        mem_id = mems['ID'].iloc[imem]
        if mem_id == -123456:
            nmem = imem
        else:
            new_mems.append(mem_id)

    grups = pd.read_excel(xlpath, sheet_name='FLOOD', header=None, skiprows=1, usecols='A:B')
    grups.fillna(-123456, inplace=True)
    ngrup = len(grups.index)
    if grups[1].iloc[1] == -123456:
        ngrup = 0
    grups_to_flood = []
    hydro_str = ''
    if ngrup > 0:
        water_depth = grups.iloc[0][1]
        irow = 0
        for row in grups.itertuples():
            if irow > 0:
                grups_to_flood.append(row._2)
            irow += 1
        hydro_str = 'HYDRO +ZISEXTFLNO  I20.000              2.000     1.025     1.000     0.250\n'
        hydro_str += 'HYDRO2    0.900IN0.8002.000\n'

    lcsel = pd.read_excel(xlpath, sheet_name='LCSEL', header=None, skiprows=1, usecols='A:B')
    lc_type = lcsel.iloc[0][1]
    irow = 0
    lcsel_str = f'LCSEL {lc_type: <9}'
    for row in lcsel.itertuples():
        if irow == 0:
            irow += 1
            continue
        lcsel_str += f'{row._2: >5}'
        if irow % 12 == 0:
            lcsel_str += f'\nLCSEL {lc_type: <9}'
        irow += 1
    lcsel_str += '\n'

    loadcns_df = pd.read_excel(xlpath, sheet_name='LOADCN')
    loadcns = []
    loadcn_str = ''
    loadfiles = []
    for row in loadcns_df.itertuples():
        if str(row.Description)[-4:] == '.txt':
            if str(row.Keep_YN).lower() == 'y':
                loadfiles.append(row.Description)
                loadcn_str += f'****   LOAD CASE:  {row.ID: <66}*\n'
                continue
        if isinstance(row.ID, str):
            ldcn_name = f'{row.ID: >4}'
        else:
            ldcn_name = f'{int(row.ID):0>4}'
        if str(row.Keep_YN).lower() == 'y':
            loadcns.append(ldcn_name)
            loadcn_str += f'****   LOAD CASE:  {ldcn_name} = {row.Description: <59}*\n'

    lcomb = pd.read_excel(xlpath, sheet_name='LCOMB', skiprows=1)

    notes = pd.read_excel(xlpath, sheet_name='Notes', header=None, skiprows=1)

    # First, loop through file to find indices where new lines will be inserted
    vars = ['MEMBER', 'JOINT', 'CODE', 'LOAD']
    ends = []
    sections = {}
    with open(PATH.joinpath(basename), 'r') as f:
        iln = 0
        for line in f:
            for var in vars:
                if line[:len(var)] == var:
                    sections[var] = iln
                if line.strip() == 'END':
                    ends.append(iln)
            iln += 1
    sections['END'] = ends[0]

    copy_line = True
    outstr = ''
    with open(PATH.joinpath(basename), 'r') as f:

        iln = -1
        for line in f:
            iln += 1

            if line[:5] == 'LDOPT':
                if ngrup > 0:  # if FLOODED
                    outstr += line[:41] + f'{water_depth:7.2f}'
                    outstr += 'GLOBMN    HYD   CMB\n'
                else:
                    outstr += line
                continue

            if line[:4] == 'GRUP' and len(line) > 6:
                grup_label = line[5:8].strip()
                if grup_label in grups_to_flood:
                    newline = line[:69] + 'F' + line[70:]
                    outstr += newline
                    continue

            if line.strip() == 'TITLE':
                outstr += f'     {title}\n'
                continue

            if 'BASIC LOAD CASES' in line:
                outstr += line + loadcn_str
                continue

            if line.strip() == '***ADD NOTES':
                for row in notes.itertuples():
                    outstr += str(row._1) + '\n'
                continue

            if line.strip() == '***ADD LOADS':
                if len(loadfiles) > 0:
                    outstr += read_loadfiles(loadfiles) + '\n'
                continue

            if line.strip() == '***ADD LCOMB':
                outstr += lcomb_str(lcomb) + 'END\n'
                continue

            if line.strip() == '***ADD CDM MGROV PGROV' and ngrup > 0:
                fpath = PATH.joinpath('insert_files', 'inplace_MGROV CDM.txt')
                with open(fpath, 'r') as f:
                    for line in f:
                        outstr += line
                continue

            if 'ANALYSIS TYPE' in line:
                outstr += f'****   ANALYSIS TYPE  : {analysis_type: <61}*\n'
                continue

            # LCSEL is inserted after last CODE
            if iln == sections['CODE'] + 1:
                outstr += lcsel_str + line
                continue

            # First END is ignored, moved to end of LCOMB
            if iln == sections['END']:
                continue

            # HYDRO strings are inserted before UCPART
            if line[:6] == 'UCPART' and ngrup > 0:
                outstr += hydro_str + line
                continue

            # Modify existing joint
            if line[:5] == 'JOINT' and len(line) > 10:
                jnt_id = line[6:10]
                if jnt_id in jnts['JNT'].values:
                    if jnt_id in new_joints:
                        new_joints.remove(jnt_id)
                    else:
                        # This skips the second line (i.e. fixity) of an existing joint
                        continue
                    ind = jnts.index[jnts['JNT'] == jnt_id].tolist()
                    jnt_dict = jnts.iloc[ind].to_dict(orient='records')[0]
                    newline = jnt_coords(jnt_dict)
                    if newline == 'xxx': # spreadsheet cells were blank
                        outstr += jnt_str(jnt_dict, line.strip()) + '\n'
                    else:
                        outstr += jnt_str(jnt_dict, newline) + '\n'
                    continue

            # Add new joints at end of JOINT section
            if iln == sections['JOINT'] + 1:
                if len(new_joints) > 0:
                    for jnt_id in new_joints:
                        ind = jnts.index[jnts['JNT'] == jnt_id].tolist()
                        jnt_dict = jnts.iloc[ind].to_dict(orient='records')[0]
                        newline = jnt_coords(jnt_dict)
                        outstr += jnt_str(jnt_dict, newline) + '\n'
                outstr += line
                continue

            # Modify existing member
            if line[:6] == 'MEMBER' and len(line) > 10:
                mem_id = line[7:15]
                if mem_id in mems['ID'].values:
                    if mem_id in new_mems:
                        new_mems.remove(mem_id)
                    else:
                        # This skips the second line (i.e. offset) of an existing member
                        continue
                    ind = mems.index[mems['ID'] == mem_id].tolist()
                    mem_dict = mems.iloc[ind].to_dict(orient='records')[0]
                    outstr += mem_str(mem_dict, line.strip()) + '\n'
                    continue

            # Add new members
            if (iln == sections['MEMBER'] + 1) and nmem > 0:
                if len(new_mems) > 0:
                    for mem_id in new_mems:
                        ind = mems.index[mems['ID'] == mem_id].tolist()
                        mem_dict = mems.iloc[ind].to_dict(orient='records')[0]
                        outstr += mem_str(mem_dict, ' ') + '\n'
                outstr += line
                continue

            # Stop copying if loadcn is not in list of loadcns to keep
            if line[:6] == 'LOADCN':
                ldcn_name = line[6:10]
                if ldcn_name not in loadcns:
                    copy_line = False
                else:
                    copy_line = True

            if iln == sections['LOAD']:
                copy_line = True

            if copy_line:
                outstr += line

    with open(PATH.joinpath(newname), 'w') as f:
        f.write(outstr)


def main():

    # xlname = 'TransportModel.xlsx'
    xlname = 'LiftModel.xlsx'
    basename = 'sacinp.PLEM_basemodel_v8-pipingremvd'
    newname = 'sacinp.PLEM_Testmodel'

    make_new_model(xlname, basename, newname)


if __name__ == "__main__":
    main()
