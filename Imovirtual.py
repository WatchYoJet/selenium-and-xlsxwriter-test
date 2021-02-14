from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter
from selenium.webdriver.common.keys import Keys
from os import system

system('cls')

print(
    'Bem Vindo ao programa de apoio à pesquisa!\nPor favor, escreva o que pretende pesquisar no Imovirtual.'
)
sn = input('\033[1;36;40mA sua pesquisa é uma freguesia? [S/N]\n\033[0m')
if sn == 's' or sn == 'S':
    system('cls')
    print(
        '\nPor favor, ao digitar, deve:\n -Pesquisar sem acentos ou "ç" \n -Colocar o concelho à frente'
    )
    search = input('\033[1;36;40m Pesquisa: \033[0m')
else:
    system('cls')
    print('\nPor favor, ao digitar, deve:\n -Pesquisar sem acentos ou "ç"')
    search = input('\033[1;36;40m Pesquisa: \033[0m')

pote = ''
for i in search:
    if i == ' ':
        pote += '-'
    else:
        pote += i.lower()

search = pote
system('cls')
tipodic = {
    '1': 'apartamento',
    '2': 'moradia',
    '3': 'quarto',
    '4': 'terreno',
    '5': 'loja',
    '6': 'armazem',
    '7': 'garagem',
    '8': 'escritorio',
    '9': 'predio',
    '10': 'quintaeherdade',
    '11': 'trespasse',
    '12': ''
}
print('\nDigite o tipo de anúncio:\
    \n1-Apartamentos\
    \n2-Moradias\
    \n3-Quartos\
    \n4-Terrenos\
    \n5-Lojas\
    \n6-Armazéns\
    \n7-Garagens\
    \n8-Escritórios\
    \n9-Prédios\
    \n10-Quintas e Herdades\
    \n11-Trespasses\
    \n12-Todos os anúncios')
tipo = input('Escolha:')
tipo = tipodic[tipo]

system('cls')
print('Digite o tipo:')
if tipo == 'apartamento' or tipo == 'apartamento' or tipo == '':
    print('1-arrendar\n2-comprar\n3-ferias')
    compra = input('Escolha:')
elif tipo == 'quarto' or tipo == 'trespasse':
    print('1-arrendar')
    compra = input('Escolha:')
else:
    print('1-arrendar\n2-comprar')
    compra = input('Escolha:')
compradic = {'1': 'arrendar', '2': 'comprar', '3': 'ferias'}
compra = compradic[compra]

system('cls')

asso = ''
if tipo == 'apartamento' or tipo == 'moradia' or tipo == '':
    print(
        '0-T0\n1-T1\n2-T2\n3-T3\n4-T4\n5-T5\n6-T6\n7-T7\n8-T8\n9-T9\n10-T10 ou superior'
    )
    asso = input('Escolha:')

asoodic = {
    '': '',
    '1': '1',
    '2': '2',
    '3': '3',
    '4': '4',
    '5': '5',
    '6': '6',
    '7': '7',
    '8': '8',
    '9': '9',
    '10': '10',
    '0': 'zero'
}
asso = asoodic[asso]

if asso and not tipo:
    url = f'https://www.imovirtual.com/{compra}/{search}/?search%5Bfilter_enum_rooms_num%5D%5B0%5D={asso}&nrAdsPerPage=72'
elif not asso and tipo:
    url = f'https://www.imovirtual.com/{compra}/{tipo}/{search}/?&nrAdsPerPage=72'
else:
    url = f'https://www.imovirtual.com/{compra}/{tipo}/{search}/?search%5Bfilter_enum_rooms_num%5D%5B0%5D={asso}&nrAdsPerPage=72'

path = "C:\Program Files (x86)\chromedriver.exe"

# start web browser
driver = webdriver.Chrome(path)

driver.get(url)

dicPote = {1: 'A', 2: 'B', 3: 'C', 4: 'D'}
pote = 1
outWorkbook = xlsxwriter.Workbook(f'Imovirtual Analise {search}.xlsx')
outSheet = outWorkbook.add_worksheet()

if asso and compra != 'arrendar':
    outSheet.write('A1', "Tipologia")
    outSheet.write('B1', "Preço")
    outSheet.write('C1', "m²")
    outSheet.write('D1', "€/m²")
    outSheet.write('E1', "Link")
elif compra == 'arrendar':
    outSheet.write('A1', "Tipologia")
    outSheet.write('B1', "Preço/mes")
    outSheet.write('C1', "m²")
    outSheet.write('D1', "Link")
else:
    outSheet.write('A1', "Preço")
    outSheet.write('B1', "m²")
    outSheet.write('C1', "€/m²")
    outSheet.write('D1', "Link")
count1 = 2
count2 = 1

try:
    main = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "body-container")))
    articles = main.find_elements_by_tag_name("article")
    for article in articles:
        x = article.get_attribute("data-url")
        details = article.find_elements_by_class_name("params")
        for detail in details:
            finals = detail.find_elements_by_tag_name("li")
            for final in finals:
                if asso:
                    if count2 != 1 and str(final.text)[0] == 'T':
                        count1 += 1
                        count2 = 1
                        outSheet.write(
                            f"{dicPote[count2]}{count1}", final.text.strip("m²€/mês"))
                    else:
                        outSheet.write(
                            f"{dicPote[count2]}{count1}", final.text.strip("m²€/mês"))
                    if count2 == 4:
                        outSheet.write(f"E{count1}", x)
                        count1 += 1
                        count2 = 1
                    else:
                        count2 += 1
                else:
                    if str(final.text)[0] != 'T':
                        if count2 != 1 and str(final.text)[0] == 'T':
                            count1 += 1
                            count2 = 1
                            outSheet.write(
                                f"{dicPote[count2]}{count1}", final.text.strip("m²€/mês"))
                        else:
                            outSheet.write(
                                f"{dicPote[count2]}{count1}", final.text.strip("m²€/mês"))
                    if count2 == 3:
                        outSheet.write(f"D{count1}", x)
                        count1 += 1
                        count2 = 1
                    else:
                        count2 += 1
finally:
    driver.quit()
    outWorkbook.close()

system('cls')

print('\n\033[1;32;40mURL:\033[0m', url)
print('\nPor favor, apenas desligue o programa depois de confirmar o exel!')
print('Se for verificado um erro, carregue "0", e contacte o developer!')
y = input('\033[1;31;40mPress 0 for logs\033[0m or any other key to close...')

if y == '0':
    if asso:
        print(sn, search, tipo, compra, asso)
        input('\033[1;31;40mPress any key to close...\033[0m')
    else:
        print(sn, search, tipo, compra)
        input('\033[1;31;40mPress any key to close...\033[0m')
