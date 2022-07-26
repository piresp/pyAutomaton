#I know isn't a good practice of programming mix Portuguese and English but my final user is from Brazil,
#so only text and the excel language is in Portuguese because I've develop on a pt-br version
#The following code take name, cpf, ddd and cellphone, format all the raw data and split in separated new tables,
#and the best part u can control it from the inputs! just say where the data is! and finally u can lay on the chair and enjoy the process (:
import pyautogui as py
from time import sleep as sleep

formuleNumb = '=SE(EXT.TEXTO(D1;3;2)<"30";CONCATENAR(ESQUERDA(D1;4);"9";DIREITA(D1;8));CONCATENAR(ESQUERDA(D1;4);DIREITA(D1;8)))'
concatNumb = '=CONCAT(D1;E1;F1)'
concatNameCPF = '=CONCAT(B1;" | CPF=";C1)'
internationalCode = '55'

tablePlaces =	{'name': str(input('Posicao da coluna NOME: ')),
                'cpf': str(input('Posicao da coluna CPF: ')),
                'ddd': str(input('Posicao da coluna DDD: ')),
                'fone': str(input('Posicao da coluna TELEFONE: '))}

archiveLenght = input('Numero total de linhas do ARQUIVO ORIGINAL: ')
cutLenght = int(input('Numero de linhas dos ARQUIVOS CRIADOS: '))
cutPieces = int(input('Numero de ARQUIVOS CRIADOS: '))

startsIn = int(input('Valor de I: '))

py.PAUSE = 0.3

def start():  
    sleep(3)
    x, y = py.position()
    screen = py.onScreen(x, y)
    if (screen == False):
        py.alert(text='MOUSE FORA DA TELA PRINCIPAL', title='ALERTA!', button='OK')
        exit()
        
    sleep(3)
    py.click()
    py.hotkey('ctrl', 'home')

def serch(here, there):
    py.hotkey('f5')
    py.write(f'{here}:{there}')
    py.hotkey('enter')

def newInstance(place):
    serch(place, place)
    py.hotkey('ctrl','+')

def replace(cut, paste):
    serch(cut, cut)
    py.hotkey('ctrl', 'x')
    newInstance(paste)

def delete(here, there):
    serch(here, there)
    py.hotkey('ctrl','-')
    py.hotkey('home')

def fill(value, here, there):
    py.write(value)
    py.hotkey('enter')
    serch(here, there)
    py.hotkey('ctrl', 'd')

def formatTable(column, value, here, there):
    newInstance(column)
    fill(value, here, there)
    
def bodyConvToValue():
    py.hotkey('ctrl', 'c')
    py.hotkey('ctrl', 'alt', 'v')
    py.hotkey('down')
    py.hotkey('down')
    py.hotkey('enter')

def convToValue(column1, column2):
    i = 0
    while i < 2:
        if i == 0:
            serch(column1, column1)
            bodyConvToValue()
        if i == 1:
            serch(column2, column2)
            bodyConvToValue()
        i += 1

def sort(args):
    for key, value in args.items():
        if (key == 'name' and value != 'A'):
            replace(value, 'A')
        if (key == 'cpf' and value != 'B'):
            replace(value, 'B')
        if (key == 'ddd' and value != 'C'):
            replace(value, 'C')
        if (key == 'fone' and value != 'D'):
            replace(value, 'D')

def save():
    i = 0
    while i < cutPieces:
        py.PAUSE = 0.6
        serch('A1',f'B{cutLenght}')
        py.hotkey('ctrlleft', 'x')
        py.hotkey('alt', 'a', 'n', 'l', interval= 0.75) 
        py.hotkey('ctrl', 'v')
        
        newInstance('1')
        py.write('First Name')
        py.hotkey('tab')
        py.write('Mobile Phone')
        py.hotkey('tab')
        
        py.hotkey('ctrl', 'a')
        py.hotkey('alt', 'a', 'c', 'y', '2', interval= 0.8)
        py.keyUp('alt')
        
        py.write(f'Pasta {i + startsIn}')
        py.hotkey('enter')
        py.hotkey('ctrl', 'w')
        
        delete('A1',f'B{cutLenght}')
        py.hotkey('down')
        i += 1
    py.hotkey('enter')
    py.alert(text='OPERAÇÃO CONCLUIDA', title='ALERTA!', button='OK')

start()
sort(tablePlaces)
delete('E','CK')
formatTable('C', internationalCode, 'C1', f'C{archiveLenght}')
formatTable('C', concatNumb, 'C1', f'C{archiveLenght}')
formatTable('C', formuleNumb, 'C1', f'C{archiveLenght}')
formatTable('A', concatNameCPF, 'A1', f'A{archiveLenght}')
delete('1','1')
convToValue('A', 'D')
delete('B', 'C')
delete('C', 'F')
save()
