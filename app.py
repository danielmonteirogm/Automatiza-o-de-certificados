from PIL import Image, ImageDraw, ImageFont
import openpyxl

# Abrir a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')

# Acessar pagina
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):

    # selecionando colunas
    nome_curso = (linha[0].value)
    nome_participante = (linha[1].value)
    tipo_de_participação = (linha[2].value)
    data_inicio = (linha[3].value)
    data_final = (linha[4].value)
    carga_horaria = (linha[5].value)
    data_emissao = (linha[6].value)

    # Escolhendo a fonte
    font_nome = ImageFont.truetype('./TAHOMABD.ttf', 90)
    font_geral = ImageFont.truetype('./TAHOMA.ttf', 80)
    fonte_data = ImageFont.truetype('./TAHOMA.TTF', 50)

    # Abrindo certificado exemplo
    image = Image.open('./certificado_padrao.jpg')

    # Editando certificado

    desenhar = ImageDraw.Draw(image)

    # nome do participante
    desenhar.text((1020, 827), nome_participante, fill='black', font=font_nome)

    # nome do curso
    desenhar.text((1060, 950), nome_curso, fill='black', font=font_geral)

    # tipo de participante
    desenhar.text((1435, 1065), tipo_de_participação,
                  fill='black', font=font_geral)

    # carga horaria
    desenhar.text((1480, 1182), str(carga_horaria) +
                  'h', fill='black', font=font_geral)

    # data de inicio
    desenhar.text((750, 1770), data_inicio, fill='blue', font=fonte_data)

    # data final
    desenhar.text((750, 1930), data_final, fill='blue', font=fonte_data)

    # data emissão
    desenhar.text((2220, 1930), data_emissao, fill='blue', font=fonte_data)

    # salvando e nomeando certificados
    image.save(f"./Certificados/{nome_participante}_certificado.png")


#


# video:https://www.youtube.com/watch?v=VwYqakOB4ow
