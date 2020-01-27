#-*- using:utf-8 -*-
import xlwings as xw
import pprint
import time


# 一つのブックの複数のシートデータを取得して配列で返す
def get_sheet_data(book, sheets):
    ret = []
    wb = xw.Book(book)

    for s in sheets:
        sheet = wb.sheets[s['sheet']]
        data = sheet.range(
            (s['range']['start_col'], s['range']['start_row']),
            (s['range']['end_col'], s['range']['end_row'])
        ).value
        ret.append(data)

    wb.close()
    wb.app.quit()

    return ret


def main():
    start_time = time.time()  # 処理時間測定開始

    book = './data.xlsx'
    sheets = \
        [
            {
                'sheet': 'param',
                'range': {
                    'start_col': 2,
                    'start_row': 1,
                    'end_col': 10,
                    'end_row': 10
                }
            },
            {
                'sheet': 'menu',
                'range': {
                    'start_col': 2,
                    'start_row': 1,
                    'end_col': 20,
                    'end_row': 4
                }
            }
        ]
    ret = get_sheet_data(book, sheets)
    param_data = ret[0]
    menu_data = ret[1]

    # variable
    variables = Variable()
    for d in param_data:
        if d[0] is None:
            break
        my_variable = VariableItem(
            d[0],
            d[1],
            d[2],
            d[3],
            d[4],
            d[5],
            d[6],
            d[7],
            d[8],
            d[9]
        )
        # pprint.pprint(my_variable)
        variables.add(my_variable)
    # variables.generate_xml()

    # menu
    menu_collection = MenuCollection()
    for m in menu_data:
        if m[1] is None:
            break
        if menu_collection.search(m[1]):
            menu_collection.add(Menu(m[1]))
        # 親があれば
        if m[0] is not None:
            for menu in menu_collection.collection:
                # 親が一致すれば親のrefに追加する
                if menu.id == m[0]:
                    menu_ref = MenuRef(m[1], m[2], m[3])
                    menu.menu_ref.append(menu_ref)
                    break

    pprint.pprint(menu_collection)
    menu_collection.generate_xml()

    end_time = time.time()
    print(end_time - start_time)

    # feature = Feature(False, True, False)
    # device_function = DeviceFunction(feature, variables)
    # device_function.generate_xml()


class VariableItem:
    def __init__(self, index, name, v_id, access_right, default_val, lower_val, upper_val,
                 single_value, v_data_type, bit_length):
        self.index: int = int(index) if index is not None else index
        self.name = name
        self.v_id = v_id
        self.access_right = access_right
        self.default_val = default_val
        self.lower_val: int = lower_val if default_val is not None else lower_val
        self.upper_val: int = upper_val if default_val is not None else upper_val
        self.single_values = single_value.splitlines() if single_value is not None else []
        self.v_data_type = v_data_type
        self.bit_length = int(bit_length) if bit_length is not None else 0

    def __repr__(self):
        data = []
        for key, value in self.__dict__.items():
            data.append(key + ':' + str(value))

        result = 'Class:' + self.__class__.__name__ + ' ('
        result += ','.join(data) + ')'
        return result

    def exists_value_range(self) -> bool:
        if self.lower_val is None:
            return False
        return True

    def exists_single_values(self) -> bool:
        if self.single_values:
            return True
        return False

    def generate_xml(self):
        print('    <Variable id="{0}" index="{1:d}" accessRights="{2:s}" defaultValue="{3}">'.format(
            self.v_id or "", self.index or "", self.access_right or "", self.default_val or ""))
        print('      <Datatype xsi:type="{}" bitLength="{}">'.format(
            self.v_data_type or "", self.bit_length
        ))

        if self.exists_value_range():
            print('        <ValueRange lowerValue="0" upperValue="50" />')
            print('      </Datatype>')

        if self.exists_single_values():
            for single_value in self.single_values:
                print('        <SingleValue value="{}">'.
                      format(single_value))
                print('          <Name textId="TI_{0}_SV_{1}" />'.
                      format(single_value, self.name))
                print('        </SingleValue>')

        print('      <Name textId="TI_{}_Name" />'.format(self.name))
        print('      <Description textId="TI_{}_Description" />'.format(self.name))
        print('    </Variable>')


class Variable:
    def __init__(self):
        self.variables = []

    def add(self, variable_item: VariableItem):
        self.variables.append(variable_item)

    def generate_xml(self):
        print("  <VariableCollection>")
        for v in self.variables:
            v.generate_xml()
        print("  </VariableCollection>")


class MenuItem:
    def __init__(self, text_id, v_ref_id):
        self.text_id = text_id
        self.v_ref_id = v_ref_id


class Feature:
    def __init__(self, block_parameter: bool, data_storage: bool, profile_characteristic: bool):
        self.block_parameter = block_parameter
        self.data_storage = data_storage
        self.profile_characteristic = profile_characteristic

    def generate_xml(self):
        print('  < Features blockParameter = "{}" dataStorage = "true" {}>'.
              format('true' if self.block_parameter else 'false',
                     'true' if self.data_storage else 'false',
                     'profileCharacteristic = "16384 32778"' if self.profile_characteristic else ""))
        print('    < SupportedAccessLocks localUserInterface = "false" '
              'dataStorage = "true" parameter = "false" localParameterization = "true" / >')
        print('  < / Features >')


class DeviceFunction:
    def __init__(self, feature: Feature, variable: Variable):
        self.feature = feature
        self.variable = variable

    def generate_xml(self):
        print('<DeviceFunction>')
        self.feature.generate_xml()
        self.variable.generate_xml()
        print('</DeviceFunction>')


class VariableRef:
    def __init__(self, variable_id):
        self.variable_id = variable_id


class MenuRef:
    def __init__(self, menu_id, condition_var, condition_val):
        self.menu_id = menu_id
        self.condition_var = condition_var
        self.condition_val = condition_val

    def generate_xml(self):
        if self.condition_var is None:
            print('  <MenuRef menuId = "M_MR_{0}" />'.format(self.menu_id))
        else:
            print('  <MenuRef menuId = "M_MR_{0}">'.format(self.menu_id))
            print('    <Condition variableId = "{0}" value = "{1}" / >'.format(
                self.condition_var, self.condition_val
            ))
            print('  </MenuRef>')


class Menu:
    def __init__(self, menu_id):
        self.id = menu_id
        self.text_id = ""
        self.variable_ref = []  # type: List[Variable]
        self.menu_ref = []  # type: List[MenuRef]

    def __repr__(self):
        data = []
        for key, value in self.__dict__.items():
            data.append(key + ':' + str(value))

        result = 'Class:' + self.__class__.__name__ + ' ('
        result += ','.join(data) + ')'
        return result

    def generate_xml(self):
        print('<Menu menuId="{0}>'.format(self.id))
        if self.menu_ref is not None:
            [mr.generate_xml() for mr in self.menu_ref]
        if self.variable_ref is not None:
            [vr.generate_xml() for vr in self.variable_ref]
        print('</Menu>')


class MenuCollection:
    def __init__(self):
        self.collection = []  # type: List[Menu]

    def search(self, menu_id):
        for c in self.collection:
            if c.id == menu_id:
                return True
        return False

    def add(self, menu: Menu):
        self.collection.append(menu)

    def generate_xml(self):
        for c in self.collection:
            c.generate_xml()


if __name__ == '__main__':
    main()
