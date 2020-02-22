from PIL import Image
import argparse
import openpyxl
import os

# Configuracao dos argumentos
parser = argparse.ArgumentParser(description = 'Software to draw images on spreadsheets.')
parser.add_argument('-i', action = 'store', dest = 'image_file_path',
                    default = '', required = True,
                    help = 'Image to be redone on a spreadsheet.')
parser.add_argument('-s', action = 'store', dest = 'resolution', required = True,
                    help = 'Image\'s resolution on the spreadsheet')

# Retorna a string hexadecimal correspondente a cor RGB informada
def rgb2hex(cor):
    return "{:02x}{:02x}{:02x}".format(cor[0], cor[1], cor[2])

def main():

    # Recebe os argumentos. Se as variaveis nao forem passadas, retorna -h
    arguments = parser.parse_args()
    imagem_original = Image.open(arguments.image_file_path)
    resolucao_imagem = int(arguments.resolution)

    try:
        # Define o nome base dos arquivos a serem criados
        nome_arquivo = os.path.splitext(os.path.basename(arguments.image_file_path))[0] + str(resolucao_imagem) + 'x' + str(resolucao_imagem)
        # Cria uma imagem reduzida que é nada mais que a imagem_original redimensionada com a resolução informada
        imagem_reduzida = imagem_original.resize((resolucao_imagem, resolucao_imagem), Image.BILINEAR)
        # Os dados da imagem sao lidos e armazenados em 'pix' 
        pix = imagem_reduzida.load()
        # Armazena as dimensoes da imagem
        largura = imagem_reduzida.size[0]
        altura = imagem_reduzida.size[1]
        # Instancia objetos para manipulacao da planilha
        wb = openpyxl.Workbook()
        ws = wb.worksheets[0]

        for x in range(1, largura + 1):
            # Altera a altura da linha
            # PARA ALTERAÇOES DO TAMANHO, MODIFICAR APENAS O VALOR QUE MULTIPLICA 18.00 PARA MANTER AS CELULAS QUADRADAS
            ws.row_dimensions[x].height = 18.00 * 1
            for y in range(1, altura + 1):
                # Altera a largura da coluna
                # PARA ALTERAÇOES DO TAMANHO, MODIFICAR APENAS O VALOR QUE MULTIPLICA 2.43 PARA MANTER AS CELULAS QUADRADAS
                ws.column_dimensions[ws.cell(x, y).column_letter].width = 2.43 * 1
                # Identifica a cor do pixel da iteracao
                cor = rgb2hex(pix[x - 1, y - 1]).upper()
                # Pinta a celula da planilha
                ws.cell(y, x).fill = openpyxl.styles.PatternFill(fgColor=cor, fill_type='solid')

        # Cria o diretorio de arquivos prontos caso não exista
        if not os.path.exists('Files_Done'):
            os.mkdir('Files_Done')

        # Salva a pixel art
        wb.save('./Files_Done/' + nome_arquivo + '.xlsx')

    except():
        print('An error occurred.')

# Chama a função main
if __name__ == '__main__':
    main()