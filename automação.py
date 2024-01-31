import openpyxl
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime 

# Carregue a planilha
workbook_pacientes = openpyxl.load_workbook('planilha-controle-estetica-receita.xlsx')
sheet_pacientes = workbook_pacientes['Plan1']

# Carregue a fonte
fonte_geral = ImageFont.truetype('tahoma.ttf', size=40)  # Ajuste o tamanho conforme necessário

# Itere sobre as linhas da planilha
for indice, linha in enumerate(sheet_pacientes.iter_rows(min_row=2), start=1):
    Nome_paciente = linha[0].value
    Procedimento = linha[1].value
    Sessão = linha[2].value
    Data = linha[3].value
    Região = linha[4].value
    Mescla = linha[5].value
    Quantidade_utilizada = linha[6].value
    Sessões_totais = linha[7].value
    Preço = linha[8].value

    # Verifica se a Data é um objeto datetime antes de tentar manipular
    if isinstance(Data, datetime):
        Data = Data.date()

    # Carregue a imagem de base
    imagem_base = Image.open('Receita-estética.jpg')

    # Crie um objeto ImageDraw
    desenho = ImageDraw.Draw(imagem_base)

    # Desenhe o texto na imagem
    desenho.text((250, 555), f'{Nome_paciente}', fill='black', font=fonte_geral)
    desenho.text((410, 660), f' {Procedimento}', fill='black', font=fonte_geral)
    desenho.text((250, 765), f' {Sessão}', fill='black', font=fonte_geral)
    desenho.text((593, 870), f' {Data}', fill='black', font=fonte_geral)
    desenho.text((535, 972), f' {Região}', fill='black', font=fonte_geral)
    desenho.text((460, 1079), f' {Mescla}', fill='black', font=fonte_geral)
    desenho.text((824, 1183), f' {Quantidade_utilizada}', fill='black', font=fonte_geral)
    desenho.text((635, 1288), f' {Sessões_totais}', fill='black', font=fonte_geral)
    desenho.text((235, 1395), f' {Preço}', fill='black', font=fonte_geral)

    # Adicione mais linhas conforme necessário

    # Salve a imagem
    imagem_base.save(f'{indice}_{Nome_paciente}_certificado.png')

# Feche a planilha
workbook_pacientes.close()
