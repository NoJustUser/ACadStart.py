import os
import subprocess
import time
import ACadAdmDist

import _ctypes
from pyautocad import Autocad, APoint

acad = Autocad()


# C:\\Users\\HP\\OneDrive\\Desktop\\BASE_
# C:\\Users\\User\\Desktop\\BASE
def generate_a_file(data_obj, file_name, dist, r_var):
    name = file_name + '.dwg'
    name_temp = file_name + '.bak'
    if r_var:
        name_folder = ACadAdmDist.make_name_folder(dist)
    else:
        name_folder = ACadAdmDist.make_name_folder_v(dist)
    folder_path = 'C:\\Users\\HP\\OneDrive\\Desktop\\BASE_\\Москва_\\' if r_var else \
        'C:\\Users\\HP\\OneDrive\\Desktop\\BASE_\\МО_\\'

    if not os.path.exists(f'{folder_path}{name_folder}'):
        # if the demo_folder directory is not present
        # then create it.
        os.makedirs(f'{folder_path}{name_folder}')
    path = f'{folder_path}{name_folder}\\{name}'
    if r_var:
        os.popen(f'copy auto_city.dwg {path}')
    else:
        os.popen(f'copy auto_village.dwg {path}')
    print('Ждите...')
    popen = subprocess.call(('start', path), shell=True)
    if popen == 1:
        while popen == 1:
            popen = subprocess.call(('start', path), shell=True)
    subprocess.Popen(f'explorer /select,{path}')
    print(popen)

    wait = 5
    while wait > 0:
        print(wait, end='')
        time.sleep(1)
        wait = wait - 1

    # Формируем вспомогательный словарь для строки с адресом
    supportive_full_adr = {f'{dist}, ': 'street_name', 'д. ': 'house_num',
                           'корп. ': 'corpus', 'строен. ': 'build_num',
                           'вл. ': 'ownership', 'кв. ': 'flat_num'}

    if popen != 1 and r_var:  # г. Москва
        # Формируем строку с адресом, в небходимом формате :
        full_adr = ''
        for k, v in supportive_full_adr.items():
            if data_obj.get(v)[-1] != '--':
                # k[:-2] здесь означает, что мы отбрасываем последине два знака строки - это - ', '
                if k[:-2] in ACadAdmDist.adm_district['Новомосковский'] + ACadAdmDist.adm_district['Троицкий']:
                    full_adr += ', '.join(data_obj.get(v)[1:]).strip() + ', '
                else:
                    full_adr += k + data_obj.get(v)[-1].strip() + ', '
        full_adr = '\\W0.97;\\T1.0;' + full_adr[0:-2]  # убираем крайнюю запятую и пробел

        # Добавляем строку с полным адресом в словарь данных объекта без учета кв.
        data_obj['full_adr'] = ','.join(full_adr.split(',')[:-1])
        # Добавляем строку с полным адресом в словарь данных объекта с учетом кв.
        data_obj['full_adr2'] = 'г. Москва, ' + full_adr
        data_obj['level2'] = data_obj['level']
        data_obj['flat_num2'] = data_obj['flat_num']
        # получаем ключи переданного в параметре функции словаря
        data_obj_keys = data_obj.keys()
        print('\n')
        print(*data_obj_keys)
        # Перезапишем значение ключа 'street_name' т. к. оно выполнило свою работу, теперь нужно
        # вывести строку формата : 'Муниципальный округ п. Сосенское '
        print(data_obj['street_name'])
        data_obj['street_name'] = ['\\W0.93;\\T0.9;', 'Муниципальный округ ', dist]
        print(data_obj['street_name'])
        # other_fields - словарь для формирования текста, который нужен для файла com_23,
        # учета сделанных работ
        other_fields = {'full_adr': 5.6, 'area': 400.0,
                        'full_adr2': 400.0, 'level2': 400.0,
                        'flat_num2': 501.0
                        }
        for v in data_obj_keys:
            color = 1 if v == 'flat_num2' else 200
            if v in other_fields.keys():
                set_text(name, text_point.get(v), ''.join(data_obj.get(v)), data_table_width.get(v), r_var,
                         other_fields.get(v))
            else:
                set_text(name, text_point.get(v), ''.join(data_obj.get(v)), data_table_width.get(v), r_var)

    elif popen != 1 and not r_var:  # Московская обл.
        other_fields_v = {'flat_num3': 500.0,
                          'full_adr': 400.0,
                          'full_adr2': 5.5,
                          'full_adr3': 5.501,
                          'level2': 5.0,
                          'flat_num2': 5.01,
                          }
        data_obj['level2'] = data_obj['level']
        data_obj['flat_num2'] = data_obj['flat_num']
        data_obj['flat_num3'] = data_obj['flat_num']
        data_obj_keys_v = data_obj.keys()
        # set_text(name, APoint(10243.38, 20686.92), 'Hello', 1500, 0, 400.0)
        for v in data_obj_keys_v:
            if v in other_fields_v:
                set_text(name, text_point_village.get(v), ''.join(data_obj.get(v)), data_table_width_village.get(v),
                         0, other_fields_v.get(v))
            else:
                set_text(name, text_point_village.get(v), ''.join(data_obj.get(v)), data_table_width_village.get(v),
                     0, 400.0)

        # Эксперимент с таблицей
        # table = acad.model.AddTable(APoint(12312.5, 20900.4), 1, 1, 7, 20)
        # table.SetTextStyle('Standard2')
        # table.TableStyle('Standard')
        # table.SetTextHeight(0, 500)
        # table.SetText(0, 0, '\T0.90;Text')
        # acad.ActiveDocument.SendCommand("_WF")
        # print(v, data_obj.get(v))
    else:
        print('Ошибка ! in Generate_File')
        try:
            acad.ActiveDocument.Application.Documents(name).Close(False, f'_{name}')
            os.popen(f'del {path} {name_temp}')
        except _ctypes.COMError:
            os.popen(f'del {path} {name_temp}')
            print(_ctypes.COMError)


def set_text(fn, p, text, width, r_var, size=4.2):
    # txt = text.split(';')[-1].strip()

    # проверяем, если передается не full_adr (полный адрес содержащий > 3-х эл-тов)
    # например, г. Москва, ул. Лескова, д. 10А, кв. 21 - это полный адрес
    # def correct_px():
    #     """Функция поправляет координату X вставляемого текста
    #        это нужно для правильного отображения текста в шапке БТИ
    #     """
    #     if len(txt.split(',')) < 3:
    #         return (len(txt) + 4) / 2
    #     return 0

    try:
        table = acad.model.AddTable(p, 1, 1, size, width)
        if r_var:
            table.StyleName = table_style.get(size, 'Standard')  # второй парамет не обязательный и
            # будет выполнен по умолч.
        else:
            table.StyleName = table_style_v.get(size, 'Standard')
        # table.StyleName = tb_style
        table.HorzCellMargin = 0
        table.VertCellMargin = 0
        table.SetCellTextHeight(0, 0, size)
        table.SetColumnWidth(0, width)

        table.SetText(0, 0, text)

    except _ctypes.COMError:
        acad.ActiveDocument.Application.Documents(fn).Close(False, f'_{fn}')
        print('Возникла непредвиденная ошибка, повторите попытку.')
        # p.x += correct_px()


table_style = {4.2: 'Standard2', 5.6: 'table_full_adr2', 400.0: 'table_full_adr2',
               500.0: 'table_full_adr2', 501: 'table_num_flat'}
table_style_v = {500.0: 'flat_num', 400.0: 'full_adr_title', 5.5: 'full_adr_center',
                 5.501: 'full_adr_left', 5.0: 'full_adr_center', 5.01: 'full_adr_center'}
text_style = {4.2: 'город_табл', 5.6: 'город_табл', 400.0: 'город_табл2'}

data_table_width = {'street_name': 140.0,
                    'street_name2': 140.0,
                    'corpus': 7,
                    'level': 7,
                    'quarter': 7,
                    'house_num': 7,
                    'ownership': 7,
                    'flat_num': 12,
                    'build_num': 7,
                    'adm_dist': 60.0,
                    'full_adr': 350.0,
                    'area': 1500.0,
                    'full_adr2': 17750.0,
                    'level2': 1500.0,
                    'flat_num2': 1500.0,
                    }

text_point = {'street_name': APoint(10243.38, 20686.92),
              'street_name2': APoint(10243.22, 20671.40),
              'corpus': APoint(10272.57, 20644.49),
              'level': APoint(10274.06, 20633.23),
              'quarter': APoint(10281.40, 20623.69),
              'house_num': APoint(10359.3, 20653.50),
              'ownership': APoint(10287.01, 20653.75),
              'flat_num': APoint(10369.09, 20633.59),
              'build_num': APoint(10368.37, 20644.90),
              'adm_dist': APoint(10297.15, 20623.7),
              'full_adr': APoint(10134.35, 20874.06),
              'area': APoint(6925.0, 23640.0),
              'full_adr2': APoint(2300.0, 25750.0),
              'level2': APoint(2800.0, 23640.0),
              'flat_num2': APoint(2800.0, 18000.0)
              }
# словарь text_point_village для области
text_point_village = {'full_adr': APoint(4547.00, 20834.56),
                      'full_adr2': APoint(10706.20, 16763.00),  # center down
                      'full_adr3': APoint(10706.20, 16810.92),  # left
                      'corpus': APoint(7290.00, 19500.23),
                      'level': APoint(5000.00, 18830.23),
                      'quarter': APoint(15281.40, 19500.23),
                      'house_num': APoint(5000.00, 19500.23),
                      'ownership': APoint(13300.00, 19500.23),
                      'flat_num': APoint(11200.00, 19500.23),
                      'build_num': APoint(9390.00, 19500.23),
                      'area': APoint(8250.0, 18830.23),
                      'level2': APoint(10769.70, 16656.50),
                      'flat_num2': APoint(10862.00, 16656.20),
                      'flat_num3': APoint(8500.00, 15656.20),
                      }

tab_style = {'street_name': 'Standard2',
             'street_name2': 'Standard2',
             'corpus': 'Standard2',
             'level': 'Standard2',
             'quarter': 'Standard2',
             'house_num': 'Standard2',
             'ownership': 'Standard2',
             'flat_num': 'Standard2',
             'build_num': 'Standard2',
             'adm_dist': 'Standard2',
             'full_adr': 'table_full_adr',
             'area': 'table_full_adr2',
             'full_adr2': 'table_full_adr2',
             'level2': 'table_full_adr2',
             'flat_num2': 'table_num_flat'
             }
# надо доделать словарь стилей таблиц для области
tab_style_village = {'full_adr': 'full_adr_title',
                     'full_adr2': 'full_adr_center',
                     'full_adr3': 'full_adr_left',
                     'corpus': 'Standard2',
                     'level': 'Standard2',
                     'quarter': 'Standard2',
                     'house_num': 'Standard2',
                     'ownership': 'Standard2',
                     'flat_num': 'Standard2',
                     'build_num': 'Standard2',
                     'area': 'Standard2',
                     'level2': 'full_adr_center',
                     'flat_num2': 'full_adr_center',
                     'flat_num3': 'flat_num',
                     }
data_table_width_village = {'full_adr': 16401.4,
                            'full_adr2': 388.0,
                            'full_adr3': 388.0,
                            'corpus': 1000,
                            'level': 1000.0,
                            'quarter': 7,
                            'house_num': 1000,
                            'ownership': 1000,
                            'flat_num': 1000,
                            'build_num': 1000,
                            'area': 1500.0,
                            'level2': 15,
                            'flat_num2': 15,
                            'flat_num3': 1000,
                            }
