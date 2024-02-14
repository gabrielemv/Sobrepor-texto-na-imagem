'''
Projeto para pegar dados de uma planilha Excel e sobrepor uma imagem

1) A planilha contém os dados dos alunos (fictícios)
2) A imagem é um certificado padrão

'''
# Pegar os dados da planilha
import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Abrir a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

# iniciar a busca à partir da linha 2
for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    #cada célula que contém a info que precisamos
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    carga_horária = linha[5].value
    
    data_inicio = linha[3].value
    data_final = linha[4].value
    
    
    data_emissao = linha[6].value
       
    # transferir os dados da planilha para a imagem do certificado
    # escolhendo a fonte da letra
    fonte_geral = ImageFont.truetype('./TAHOMA.TTF', 70)
    fonte_nome = ImageFont.truetype('./TAHOMABD.TTF', 90)
    fonte_data = ImageFont.truetype('./TAHOMA.TTF', 55)
    
    # abrir a imagem
    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)
    
    #posicionar o texto no local correto
    desenhar.text((1029,820),nome_participante,fill='black', font=fonte_nome)
    desenhar.text((1069,960),nome_curso,fill='black', font=fonte_geral)
    desenhar.text((1429,1070),tipo_participacao,fill='black', font=fonte_geral)
    desenhar.text((1477,1190),str(carga_horária),fill='black', font=fonte_geral)
    
    desenhar.text((700,1750),data_inicio,fill='black', font=fonte_geral)
    desenhar.text((700,1910),data_final,fill='black', font=fonte_geral)
    
    desenhar.text((2160,1900),data_emissao,fill='black', font=fonte_geral)
    
    
    image.save(f'./{indice} {nome_participante} certificado.png')