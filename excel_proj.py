import openpyxl
import odf
import pyautogui as py

#criar uma planilha
book = openpyxl.Workbook()
#vizualivar pagina existente
print(book.sheetnames)
#criar página com nome já especificado
book.create_sheet('Frutas')
#selecionar página
frutas_page = book['Frutas']
frutas_page.append(['Fruta', 'Quantidade', 'Preço'])
#adicionar dados
frutas_page.append(['Banana', '5', 'R$3,90'])
frutas_page.append(['Maça', '4', 'R$6'])
frutas_page.append(['Tomate', '65', 'R$3'])
#salvar planilha
book.save('Planilha de Frutas.xlsx')
