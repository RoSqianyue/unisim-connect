from interface import UniSimInterface

def main():
    # Create a new interface to UniSim
    unisim = UniSimInterface()
    
    # Connect to the active UniSim case
    unisim.open_case(use_active=True)
    
    # Print information about all material streams
    print("\nMaterial Streams Information:")
    for i in range(1, 10):  # Assuming streams are named 1, 2, 3, etc.
        try:
            stream = unisim.get_material_stream(str(i))
            print(f"\nStream {i}:")
            print(f"  Temperature: {stream.temperature:.2f} K")
            print(f"  Pressure: {stream.pressure:.2f} bar")
            print(f"  Molar Flow: {stream.molar_flow:.2f} gmole/s")
            print(f"  Components: {stream.get_component_molar_fractions()}")
        except:
            # Stream might not exist
            pass
    
    # Get a spreadsheet if it exists
    try:
        spreadsheet = unisim.get_spreadsheet("SPREADSHEET")
        print("\nSpreadsheet Values:")
        print(f"  A1: {spreadsheet.get_cell_value('A1')}")
        print(f"  B1: {spreadsheet.get_cell_value('B1')}")
    except:
        print("\nNo spreadsheet named 'SPREADSHEET' found")

if __name__ == "__main__":
    main()