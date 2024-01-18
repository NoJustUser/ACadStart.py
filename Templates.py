import os
import subprocess
import time

from pyautocad import Autocad, APoint

name = 'шЛенинградское_12кв12' + '.dwg'  # это имя будем брать из формы
path = f'D:\\Python_projects\\pyAutocad\\{name}'  # этот путь будет сформирован по данным из формы
os.popen(f'copy auto_city.dwg {path}')
popen = subprocess.Popen(('start', name), shell=True)

wait = 3
while wait > 0:
    print(wait, end='')
    time.sleep(1)
    wait = wait - 1

acad = Autocad()
acad.prompt("\nHello, Autocad from Python\n")
print(acad.doc.Name)

# p1 = APoint(0, 0)
# p2 = APoint(50, 25)
# for i in range(5):
#     text = acad.model.AddText('Hi %s!' % i, p1, 2.5)
#     acad.model.AddLine(p1, p2)
#     acad.model.AddCircle(p1, 10)
#     p1.y += 10

# dp = APoint(10, 0)
# for text in acad.iter_objects('Text'):
#     print('text: %s at: %s' % (text.TextString, text.InsertionPoint))
#     text.InsertionPoint = APoint(text.InsertionPoint) + dp

# for obj in acad.iter_objects(['Circle', 'Line']):
#     print(obj.ObjectName)

p3 = APoint(7300.0, 28633.0)
p4 = APoint(600, 350)
p5 = APoint(10297.5, 20621.7)
# text = acad.model.AddText('Hello... !', p3, 700.0)

# text = acad.model.AddMText(p3, 20000, '-- АО')
# text.Height = 400
# text.Color = 200

adm_dist = 'Северо-Восточный'  # эта переменная будет браться из словаря {'район': 'значение'}
admin_district = f'\T0.9;{adm_dist}' + ' АО'  # 0.9 здесь это межстрочный интервал
administrative_district_text = acad.model.AddMText(p5, 20000, admin_district)
administrative_district_text.Height = 4
administrative_district_text.StyleName = 'город_табл'
administrative_district_text.Color = 200

# text1 = acad.model.AddMText(APoint(10355.26, 20621.59), 2000,
#                             f'\\fVerdana|b0|i0|c204|p34;\H0.84x;\W1;\T0.9; г. Москвы')
# text1.Height = 5
# text1.StyleName = 'standard'
# text1.Color = 200

# myLine = acad.model.AddLine(p3, p4)
# myLine.Lineweight = 30
# # проверка назначения слоя
# print("текущий слой: " + str(myLine.Layer))
# # проверка текущего типа линий
# print("текущий тип линии: " + str(myLine.Linetype))
# # проверка масштаба типа линий
# print("текущий масштаб типа линии: " + str(myLine.LinetypeScale))
# # проверка текущей толщины линии
# print("текущая толщина линии: " + str(myLine.Lineweight))
# # проверка текущей толщины
# print("текущая толщина: " + str(myLine.Thickness))
# # проверка текущий материал
# print("текущий материал:" + str(myLine.Material))
# print('Длина строки :', myLine.Length)
# text.StyleName = 'Vlad_style'
# print('Style: ', text.StyleName)


# acad.prompt("\nHello, Autocad from Python\n")
# \W0.93 - коэф. сжатия-растяжения,
# # T0.9 - Трекинг (увеличение-уменьшение расст. м/у символами)
# acad.ActiveDocument.Close()                                             # Закрывает активный файл с сохранением
# acad.ActiveDocument.Application.Documents(name).Close(True, f'_{name}') # Закрывает файл с именем name без сохранения
# Если параметр True в методе Close установить в False, то файл name будет закрыт без сохранения
