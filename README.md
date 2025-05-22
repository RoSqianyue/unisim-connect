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

## License

MIT