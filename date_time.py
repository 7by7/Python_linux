from datetime import datetime
import openpyxl

workbook_aluno = openpyxl.load_workbook('Book1.xlsx')
sheet_alunos = workbook_aluno['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2,max_row=2)):
    data_inicio = linha[4].value 

data_e_hora_atuais = datetime.now()
data_inicio_format = data_inicio.strftime('%d/%m/%y %H:%M')

print(data_e_hora_em_texto)