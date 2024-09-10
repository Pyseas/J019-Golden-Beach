# golden_beach
golden_beach is a Python library for the Golden Beach project (J019) to write and update SACS files based on data contained in spreadsheets.

## Installation
Use pip to install golden_beach
```bash
pip install golden_beach
```

## Usage
```python
import golden_beach as gb

# Creates a new sacinp file using data from LiftModel spreadsheet
gb.make_new_model('LiftModel.xlsx', 'sacinp.base', 'sacinp.lift')

# Writes a new file with piping loads in SACS format
gb.write_piping_loads('PipingLoads.xlsx', 'sacinp.base', 'sacinp.lift')

# Writes a new file with soil spring data in SACS format
rng = gb.SoilRanges(tz='B10:G30', qz='I10:Q30', py='S10:AH62')

gb.write_soil_springs('Springs.xlsx', 'psi_low.dat', rng, 'LOW ESTIMATE')
```
