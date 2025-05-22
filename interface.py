import win32com.client
import yaml
import win32gui

class UniSimInterface:
    def __init__(self):
        self.app = win32com.client.Dispatch('UnisimDesign.Application')
        self.app.Visible = True
        self.doc = None
        self.flowsheet = None
        self.case = None
        self.spreadsheet = None

    def open_case(self, case_file_path=None, use_active=True, case_name=None):
        if use_active:
            if case_name:
                print(f"Trying to use active UniSim case: {case_name}")
                self.case = self.find_case_by_name(case_name)
                if self.case:
                    self.case.Activate()
                    print("Case activated successfully!")
                else:
                    print(f"Case '{case_name}' not found.")
                    return
            elif case_file_path:    
                print(f"Opening UniSim case: {case_file_path}")
                self.case = self.app.SimulationCases.Add(case_file_path)
            else:
                instances, windows = self.get_running_unisim_instances()
                print(f"Found {len(windows)} UniSim window(s) and {len(instances)} COM instance(s)")

                for i, instance in enumerate(instances):
                    try:
                        print(f"COM Instance {i+1}:")
                        print(f"  Version: {instance.Version}")
                        print(f"  Active Document: {instance.ActiveDocument.Title.Value}")
                    except:
                        print(f"  Unable to retrieve details for COM instance {i+1}")

                for i, hwnd in enumerate(windows):
                    print(f"Window {i+1}:")
                    print(f"  Title: {win32gui.GetWindowText(hwnd)}")
                    print(f"  Handle: {hwnd}")

                self.doc = self.app.ActiveDocument

        elif case_file_path:
            print(f"Opening UniSim case: {case_file_path}")
            self.case = self.app.SimulationCases.Open(case_file_path)
        else:
            print("Using current UniSim case")
            self.doc = self.app.ActiveDocument if self.app.ActiveDocument else None

        if self.case:
            self.case.Activate()
            self.doc = self.app.ActiveDocument
            self.flowsheet = self.case.Flowsheet
        else:
            print("No case file specified and no active document found.")

    def save_case(self):
        if self.case:
            print("Saving UniSim case")
            self.case.Save()
        else:
            print("No case to save")
    
    def get_stream(self, stream_name):
        return self.flowsheet.Streams.Item(stream_name)

    def get_material_stream(self, stream_name):
        return MaterialStream(self.flowsheet.MaterialStreams.Item(stream_name))

    def get_energy_stream(self, stream_name):
        return EnergyStream(self.flowsheet.EnergyStreams.Item(stream_name))

    def get_operation(self, operation_name):
        return self.flowsheet.Operations.Item(operation_name)

    def get_spreadsheet(self, spreadsheet_name):
        return Spreadsheet(self.flowsheet.Operations.Item(spreadsheet_name))

    def get_running_unisim_instances(self):
        unisim_instances = []
        unisim_windows = []

        def enum_window_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if "UniSim Design" in window_text:
                    results.append(hwnd)
            return True

        win32gui.EnumWindows(enum_window_callback, unisim_windows)

        for hwnd in unisim_windows:
            try:
                instance = win32com.client.GetActiveObject("UnisimDesign.Application")
                if instance not in unisim_instances:
                    unisim_instances.append(instance)
            except:
                print(f"Unable to get COM object for window: {win32gui.GetWindowText(hwnd)}")

        return unisim_instances, unisim_windows

    def find_case_by_name(self, case_name):
        for case in self.app.SimulationCases:
            if case.Title.Value == case_name:
                return case
        print(f"Case '{case_name}' not found.")
        return None



class MaterialStream:
    def __init__(self, unisim_stream):
        self.unisim_stream = unisim_stream

    def get_property(self, property_name, units=None):
        prop = getattr(self.unisim_stream, property_name)
        if units:
            return prop.GetValue(units)
        else:
            return prop.GetValue()

    def set_property(self, property_name, value, units=None):
        prop = getattr(self.unisim_stream, property_name)
        if units:
            prop.SetValue(value, units)
        else:
            prop.SetValue(value)

    @property
    def temperature(self):
        return self.get_property('Temperature', 'K')

    @temperature.setter
    def temperature(self, value):
        self.set_property('Temperature', value, 'K')

    @property
    def pressure(self):
        return self.get_property('Pressure', 'bar')

    @pressure.setter
    def pressure(self, value):
        self.set_property('Pressure', value, 'bar')

    @property
    def molar_flow(self):
        return self.get_property('MolarFlow', 'gmole/s')

    @molar_flow.setter
    def molar_flow(self, value):
        self.set_property('MolarFlow', value, 'gmole/s')


    @property
    def component_names(self):
        names = self.unisim_stream.ComponentMolarFraction.OffsetNames
        # 将嵌套的元组展平为一个列表
        return [name for tuple in names for name in tuple if name]

    @property
    def component_molar_fraction(self):
        return self.unisim_stream.ComponentMolarFraction.Values

    def get_component_molar_fractions(self):
        names = self.component_names
        fractions = self.component_molar_fraction
        if len(names) != len(fractions):
            print(f"Warning: Number of names ({len(names)}) does not match number of fractions ({len(fractions)})")
        return dict(zip(names, fractions))
    

    def set_component_molar_fractions(self, fractions_dict):
        """
        设置物料流中各组分的摩尔分数
        :param fractions_dict: 包含组分名称和对应摩尔分数的字典
        """
        current_names = self.component_names
        new_fractions = [0] * len(current_names)
        
        for i, name in enumerate(current_names):
            if name in fractions_dict:
                new_fractions[i] = fractions_dict[name]
            else:
                print(f"Warning: Component {name} not found in input, setting to 0")
        
        # 确保摩尔分数总和为1
        total = sum(new_fractions)
        if abs(total - 1.0) > 1e-6:  # 允许小误差
            print(f"Warning: Total mole fraction is {total}, normalizing to 1")
            new_fractions = [x/total for x in new_fractions]
        
        try:
            self.unisim_stream.ComponentMolarFraction.SetValues(new_fractions)
            print("Component molar fractions updated successfully")
        except Exception as e:
            print(f"Error setting component molar fractions: {e}")

    def get_component_molar_flow(self, unit='gmole/s'):
        """
        获取指定单位的组分摩尔流量
        :param unit: 期望的流量单位，如果不指定则使用gmole/s
        :return: 包含组分名称和对应摩尔流量的字典
        """
        names_tuple = self.unisim_stream.ComponentMolarFlow.OffsetNames
        names = [name for tuple in names_tuple for name in tuple if name]
        
        try:
            if unit:
                molar_flows = self.unisim_stream.ComponentMolarFlow.GetValues(unit)
            else:
                molar_flows = self.unisim_stream.ComponentMolarFlow.Values
        except Exception as e:
            print(f"Error getting molar flows: {e}")
            print("Falling back to default values.")
            molar_flows = self.unisim_stream.ComponentMolarFlow.Values
        
        return dict(zip(names, molar_flows))


    @property
    def mass_flow(self):
        return self.get_property('MassFlow', 'kg/h')

    @property
    def heat_flow(self):
        return self.get_property('HeatFlow', 'kJ/h')

    @property
    def vapour_fraction(self):
        return self.get_property('VapourFraction')

    @property
    def molecular_weight(self):
        return self.get_property('MolecularWeight')

    @property
    def z_factor(self):
        return self.get_property('ZFactor')

class EnergyStream:
    def __init__(self, unisim_stream):
        self.unisim_stream = unisim_stream

    @property
    def heat_flow(self):
        return self.unisim_stream.HeatFlow.GetValue('kJ/h')

    @heat_flow.setter
    def heat_flow(self, value):
        self.unisim_stream.HeatFlow.SetValue(value, 'kJ/h')

class Spreadsheet:
    # 获取spreadsheet对象
    #spreadsheet = unisim.doc.Flowsheet.Operations.Item('spreadsheet')
    def __init__(self, unisim_spreadsheet):
        self.unisim_spreadsheet = unisim_spreadsheet

    def get_cell_value(self, cell):
        return self.unisim_spreadsheet.Cell(cell).CellValue

    def set_cell_value(self, cell, value):
        self.unisim_spreadsheet.Cell(cell).CellValue = value

    def get_cell_formula(self, cell):
        return self.unisim_spreadsheet.Cell(cell).CellFormula

    def set_cell_formula(self, cell, formula):
        self.unisim_spreadsheet.Cell(cell).CellFormula = formula

def load_parameters(file_path):
    with open(file_path, 'r') as file:
        return yaml.safe_load(file)