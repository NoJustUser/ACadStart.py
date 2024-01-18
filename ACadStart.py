import ACadPy
import ACadAdmDist
from tkinter import *
import traceback

temp = {'house_num': '\\T0.9;',
        'corpus': '\\W0.93;\\T0.9;',
        'build_num': '\\T0.9;',
        'ownership': '\\T0.9;',
        'quarter': '\\W0.93;\\T0.9;',
        'flat_num': '',
        'level': '\\W0.93;\\T0.9;',
        'area': '',
        }
temp_village = {'house_num': '',
                'corpus': '',
                'build_num': '',
                'ownership': '',
                'quarter': '',
                'flat_num': '',
                'level': '',
                'area': '',
                }
data_obj = dict()


def radio_button_commands():
    if r_var.get():
        ent_adr.delete(0, END)
        ent_adr.insert(0, 'г. Москва, ')
    else:
        ent_adr.delete(0, END)
        ent_adr.insert(0, 'Московская обл., ')


def create_file():
    # из поля ввода адреса получаем название улицы/района(поселения) (st_name/dist)
    st_name = ent_adr.get().split(',')[-1].strip()
    dist = ent_adr.get().split(',')[1].strip()
    # Если Москва
    if r_var.get():
        # при этом если АО Новомосковский, то возвращаем строку типа "п. Сосенское, ул. Лесные Поляны"
        res_search_dist = ACadAdmDist.dist_search(dist)
        if res_search_dist in ['Новомосковский', 'Троицкий']:
            tmp = ent_adr.get().split(',')[1:]
            data_obj['street_name'] = ['\\W0.93;\\T0.9;', ','.join(tmp)]
            if len(tmp) == 2:
                tmp2 = tmp[-1]
                data_obj['street_name2'] = ['\\W0.93;\\T0.9;', tmp2]
                # 'п. Сосенское, ул. Лесная',
            else:  # Если len(tmp) > 2 напр.: п. Сосенское, пос. Коммунарка, ул. Лесная
                tmp2 = tmp[1:]
                data_obj['street_name2'] = ['\\W0.93;\\T0.9;', *tmp2]
                # data_obj['street_name2'] = ['\W0.93;\T0.9;', ','.join(tmp2)]
                # 'п. Сосенское, / пос. Коммунарка, ул. Лесная' - отбрасывается п. Сосенское
        else:
            data_obj['street_name'] = ['\\W0.93;\\T0.9;', st_name]
            data_obj['street_name2'] = data_obj['street_name']  # 'ул. Лесная'
        # в цикле формируем словарь data_obj для передачи данных в файл ACadPy.py
        # значение en - объект поля ввода, из которого получаем введенные данные при
        # помощи метода get()
        for en, v in ent_identities:
            data_obj[v] = [temp[v], en.get()]
        # значение округа получаем из словаря данных округов (файл ACadAdmDist.py)
        data_obj['adm_dist'] = ['\\T0.9;', res_search_dist + ' АО']
        # # формируем имя файла
        file_name = ACadAdmDist.create_filename(st_name, data_obj)
        try:
            ACadPy.generate_a_file(data_obj, file_name, dist, r_var)
            for k, v in data_obj.items():
                print(k, ' : ', v)
            print('Имя файла : ', file_name)
            print('Округ : ', dist)
            print('Город/Обл. - 1/0 : ', r_var)
        except Exception:
            for k, v in data_obj.items():
                print(k, ' : ', v)
            print('Имя файла : ', file_name)
            print('Округ : ', dist)
            print('Город/Обл. - 1/0 : ', r_var)
            print('Исключение в Tkinter')
            print(traceback.print_stack())
    # Иначе если Московская обл.
    else:
        print('village')
        # Формируем вспомогательный словарь для строки с адресом
        supp_full_adr_village = {'д. ': 'house_num', 'корп. ': 'corpus',
                                 'строен. ': 'build_num', 'вл. ': 'ownership',
                                 'кв. ': 'flat_num'}
        data_obj['full_adr'] = ent_adr.get().split(',')
        for en, v in ent_identities:
            data_obj[v] = [temp_village[v], en.get()]
        for k, v in supp_full_adr_village.items():
            if data_obj[v][1] != '--':
                data_obj['full_adr'].append(k + data_obj[v][1])
        data_obj['full_adr2'] = ', '.join(data_obj['full_adr'][0:-1])
        data_obj['full_adr3'] = ', '.join(data_obj['full_adr'][0:-1])
        data_obj['full_adr'] = ', '.join(data_obj['full_adr'])
        # получаем р-н в переменную dist_village
        dist_village = ACadAdmDist.dist_search_village(ent_adr.get())
        # # формируем имя файла
        file_name = ACadAdmDist.create_filename(st_name, data_obj)

        # Тестовый вывод data_obj:
        print('Район : ', dist_village)
        print('Имя файла : ', file_name)
        for k, v in data_obj.items():
            print(k, ' : ', v)

        try:
            ACadPy.generate_a_file(data_obj, file_name, dist_village, 0)
        except Exception:
            print('Исключение в Tkinter')


# ===================== Создание окна и компоновка элементов =============================
root = Tk()
root.title('Работа с файлами AutoCad')
w = root.winfo_screenwidth()
h = root.winfo_screenheight()
w = w // 2  # середина экрана
h = h // 2
w = w - 200  # смещение от середины
h = h - 200
root.geometry(f'616x230+{w}+{h}')
f_top = LabelFrame(text='Введите адрес объекта [р-н, п.], [пер., ул.] :')
f_bot = LabelFrame(text='Введите данные объекта :')
f_top.grid(row=0, column=0, columnspan=10, padx=10, pady=5, ipadx=0, ipady=5)
f_bot.grid(row=1, column=0, columnspan=10, padx=0, pady=0, ipadx=3, ipady=5)

# Здесь надо сформировать словарь полей lab и ent для заполнения данных в форме
# это нужно для компановки элементов полей и текста на форме
data_fields = {'lab_house_num': ['№ дома :', 'house_num'],
               'lab_corpus': ['№ корп. :', 'corpus'],
               'lab_build_num': ['№ стр. :', 'build_num'],
               'lab_ownership': ['№ влд. :', 'ownership'],
               'lab_quarter': ['№ кв-ла :', 'quarter'],
               'lab_flat_num': ['№ кв. :', 'flat_num'],
               'lab_level': ['№ этажа :', 'level'],
               'lab_area': ['площадь, м2 :', 'area']
               }
# создаем список для добавления объектов полей для индентификации
# их в случае необходимости
ent_identities = []
# Переменная для связывания радиокнопок
r_var = BooleanVar()
r_var.set(True)

city = Radiobutton(f_top, text='г. Москва', variable=r_var, value=True, command=radio_button_commands)
villiage = Radiobutton(f_top, text='Московская обл.', variable=r_var, value=False, command=radio_button_commands)

city.grid(row=0, column=0, sticky=W)
villiage.grid(row=0, column=1, sticky=W)

ent_adr = Entry(f_top, width=98)
# поле ввода адреса (г. Москва, прописываем сразу, по умолчанию)
ent_adr.insert(0, 'г. Москва, ')
# в цикле компануем поля ввода данных по объекту
row, col = 1, -1
for key, value in data_fields.items():
    col += 1
    label = Label(f_bot, text=value[0])
    ent = Entry(f_bot, width=9)
    ent.insert(0, '--')
    ent_identities.append((ent, value[1]))
    if key in ['lab_flat_num', 'lab_level']:
        col = 0
        row += 1
    label.grid(row=row, column=col, sticky=W, padx=0, pady=2)
    col += 1
    ent.grid(row=row, column=col, sticky=W, padx=0, pady=2)

btn_create = Button(text='сформировать', width=20, command=create_file)
ent_adr.grid(row=1, column=0, columnspan=9, sticky=W, padx=2, pady=2)
btn_create.grid(row=3, column=0, columnspan=10, padx=2, pady=2)

root.mainloop()
