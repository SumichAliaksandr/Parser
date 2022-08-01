import os
from datetime import datetime
import requests
from openpyxl import load_workbook, Workbook
def parsing_synapsenet(name, adres):
    """
    Парсинг ИНН с сайта https://synapsenet.ru/ по
    name - название организации
    adres - адрес организации
    Генерирует запрос по названию организации. Преобразует его в строку. В полученной строке находит организацию
    с почтовым индексом, указанным в адресе и возвращает ее ИНН, если он число, иначе - возвращает "Ошибка"
    """

    adr = adres.split()
    index = adr[0]
    now = datetime.now()
    data = now.strftime("%d.%m.%Y")

    name_splited = name.split()
    name = ''
    for na in name_splited:
        name = name + '%20' + na
    name = name[3:]

    zapros = 'https://synapsenet.ru/searchorganization/proverka-kontragentov?query=' \
             + name + '%22&startDate=01.01.1990&endDate=' \
             + data +'&regionId=%7B%22District%22:[],%22Region%22:[],%22Area%22:[],%22Loc%22:[]%7D&okved=[]'

    #print(zapros)

    r = requests.get(zapros)
    sait = r.text

    f = open('zapros.txt', 'w', encoding='UTF-8')
    f.write(sait)
    f.close()
    if (sait.find('Карточка организации') == -1) and (sait.find('Попробуйте изменить поисковый запрос') == -1):
        i = sait.find(index)

        if i > -1:
            sait = sait[:i]
            marker_inn = 'ИНН </span>'
            inn_position = sait.rfind(marker_inn) + 11
            inn = sait[inn_position:(inn_position + 10)]

        else:
            marker_inn = 'ИНН </span>'
            inn_position = sait.find(marker_inn) + 11
            inn = sait[inn_position:(inn_position + 10)]

        if inn.isdigit():
            return inn
        else:
            print('Синапс')
            return 'Ошибка'
    elif sait.find('Карточка организации') > -1:
        print('Точно нашли организацию')
        pos = sait.find('ИНН ')
        inn = sait[pos:pos+14]
        inn = inn[4:]

        if inn.isdigit():
            return inn
        else:

            return 'Ошибка'
    else:
        return 'Ошибка'



def parsing_rusprofile(name, adres):
    """
    Парсинг ИНН с сайта https://www.rusprofile.ru/ по
    name - название организации
    adres - адрес организации
    Генерирует запрос по названию организации. Преобразует его в строку. В полученной строке находит организацию
    с почтовым индексом, указанным в адресе и возвращает ее ИНН, иначе - возвращает "Ошибка"
    """
    dop_info = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36',
                'accept': '*/*'}
    adr = adres.split()
    index = adr[0]

    name_splited = name.split()
    name = ''
    for na in name_splited:
        name = name + '+' + na
    name = name[1:]
    zapros = 'https://www.rusprofile.ru/search?query='+name+'&type=ul'
    print(zapros)
    r = requests.get(zapros, dop_info)
    sait = r.text

    #f = open('zapros.txt', 'w', encoding='UTF-8')
    #f.write(sait)
    #f.close()

    i = sait.find(index)
    inn = ''
    if i > -1:
        sait = sait[i:]
        marker_inn = '<dt>ИНН</dt>'
        inn_position = sait.find(marker_inn) + 37
        inn = sait[inn_position:(inn_position + 10)]
        #print(inn)


    if inn.isdigit():
        return inn
    else:
        return 'Ошибка'
"""
organizacii = (('ООО "Ремонтник', '368730, республика Дагестан, Ахтынский район, село Ахты, улица Арсена Байрамова, дом 153'),
               ('Общество с Ограниченной Ответственностью "Алан-Тревел"', '367000, республика Дагестан, город Махачкала, улица Богатырева, 12'),
               ('ООО "Кубаньстройкреп"', '350039, Краснодарский край, город Краснодар, Майский проезд, дом 5, офис 310'),
               ('ООО "Инвестстрой"', '353901, Краснодарский край, город Новороссийск, Элеваторная улица, дом 33 литер б, офис 6'))
for org in organizacii:
    print(parsing_rusprofile(org[0], org[1]))



organizacii = (('ООО "Кубаньстройкреп"', '350039, Краснодарский край, город Краснодар, Майский проезд, дом 5, офис 310'),)
for org in organizacii:
    print(parsing_synapsenet(org[0], org[1]))

"""
def tablica():
    spisok_failov = os.listdir(path=".")

    for fail in spisok_failov:
        if (fail[-5:] == '.xlsx') and (fail!='res.xlsx'):
            return fail
    print('Поместите вашу таблицу в формате xlsx в папку с программой и перезапустите программу. Название не должно быть res.xlsx')
    return 'Oshibka'
