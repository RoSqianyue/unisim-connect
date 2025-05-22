# UniSim Connect

A Python interface for interacting with UniSim Design through COM automation.

## Features

- Connect to running UniSim Design instances
- Open, manipulate, and save UniSim cases
- Access and modify material streams properties
- Access and modify energy streams
- Work with UniSim spreadsheets

## Requirements

- Windows OS
- UniSim Design installed
- Python 3.6+
- Required Python packages (see requirements.txt)

## Installation

```bash
pip install -r requirements.txt
```

## Usage Example

### Working with Material Streams

```python
from unisim_connect.interface import UniSimInterface

# Create interface
unisim = UniSimInterface()

# Connect to an active UniSim case or open a new one
unisim.open_case(use_active=True)

# Get a material stream
stream = unisim.get_material_stream("1")

# Read properties
print(f"Temperature: {stream.temperature} K")
print(f"Pressure: {stream.pressure} bar")
print(f"Molar Flow: {stream.molar_flow} gmol/s")

# Modify properties
stream.temperature = 298.15
stream.pressure = 1.01325

# Save the case
unisim.save_case()
```

### Working with Spreadsheets

```python
from unisim_connect.interface import UniSimInterface

# Create interface and connect
unisim = UniSimInterface()
unisim.open_case(use_active=True)

# Get a spreadsheet by name
spreadsheet = unisim.get_spreadsheet("SS1")

# Read cell values
value_a1 = spreadsheet.get_cell_value("A1")
value_b2 = spreadsheet.get_cell_value("B2")
print(f"Cell A1 value: {value_a1}")
print(f"Cell B2 value: {value_b2}")

# Write values to cells
spreadsheet.set_cell_value("C1", 42.5)
spreadsheet.set_cell_value("C2", "Sample Text")

# Read cell formulas
formula = spreadsheet.get_cell_formula("D1")
print(f"Cell D1 formula: {formula}")

# Set cell formula
spreadsheet.set_cell_formula("D2", "=SUM(A1:C1)")

# Save changes
unisim.save_case()
```

## License

MIT