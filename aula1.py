"""

"""

# Pegar os dados da planilha
import openpyxl
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime

#abrir planilha
workbook_aluno = openpyxl.load_workbook('Book1.xlsx')
sheet_alunos = workbook_aluno['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2,max_row=2)):
    #acessa cada célula
    nome_curso = linha[0].value 
    nome_participante = linha[1].value 
    area = linha[2].value 
    carga_horaria = linha[3].value 
    data_inicio = linha[4].value 
    data_termino = linha[5].value
    data_emissao = linha[6].value 
    
    data_inicio_format = data_inicio.strftime('%d/%m/%y %H:%M')
    data_termino_format = data_termino.strftime('%d/%m/%y %H:%M')
    data_emissao_format = data_emissao.strftime('%d/%m/%y')

# Transferir para a imagem do certificado
font_nome = ImageFont.truetype('./VIVALDII.TTF',210)
font_geral = ImageFont.truetype('./MyriadPro-Regular.otf',40)
font_data = ImageFont.truetype('./MyriadPro-Regular.otf',25)

image = Image.open('./certificate_padrao2.png')
desenhar = ImageDraw.Draw(image)

desenhar.text((905,1005),nome_participante,fill='black',font=font_nome)
desenhar.text((2070,1710),nome_curso,fill='black',font=font_geral)
#desenhar.text((1135,1110),area,fill='black',font=font_geral)
#desenhar.text((1440,1210),str(carga_horaria),fill='black',font=font_geral)
#desenhar.text((300,800),str(data_inicio_format),fill='black',font=font_data)
#desenhar.text((300,880),str(data_termino_format),fill='black',font=font_data)
#desenhar.text((1020,880),str(data_emissao_format),fill='black',font=font_data)

image.save(f'./{indice} {nome_participante}certificado.png')