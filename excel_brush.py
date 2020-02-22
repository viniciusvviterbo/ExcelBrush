from PIL import Image
import argparse
import openpyxl

# Configuracao dos argumentos
parser = argparse.ArgumentParser(description = 'Um programa para desenhar fotos em planilhas.')
parser.add_argument('-i', action = 'store', dest = 'image',
                    default = '', required = True,
                    help = 'A imagem a ser processada pelo programa.')
parser.add_argument('-s', action = 'store', dest = 'outputImageSize', required = True,
                    help = 'O tamanho da imagem a ser retornada')

def rgb2hex(cor):
    return "{:02x}{:02x}{:02x}".format(cor[0], cor[1], cor[2])

def main():

    # Recebe os argumentos, se as variaveis nao forem passadas, retorna -h
    arguments = parser.parse_args()

    imagem_original = Image.open(arguments.image)
    resolucao_imagem = int(arguments.outputImageSize)

    try:
        # Define o nome base dos arquivos a serem criados
        nome_arquivo = 'andre_' + str(resolucao_imagem) + 'x' + str(resolucao_imagem)
        # Cria uma imagem reduzida que é nada mais que a imagem_original redimensionada com a resolução informada
        imagem_reduzida = imagem_original.resize((resolucao_imagem, resolucao_imagem), Image.BILINEAR)
        # Cria a imagem final redimensionando a imagem_reduzida ao tamanho da original
        imagem_pixelada = imagem_reduzida.resize(imagem_original.size, Image.NEAREST)

        pix = imagem_reduzida.load()

        largura = imagem_reduzida.size[0]
        altura = imagem_reduzida.size[1]

        wb = openpyxl.Workbook()
        ws = wb.worksheets[0]

        for x in range(1, largura + 1):
            ws.row_dimensions[x].height = 18.00 * 1
            for y in range(1, altura + 1):
                ws.column_dimensions[ws.cell(x, y).column_letter].width = 2.43 * 1
                cor = rgb2hex(pix[x - 1, y - 1]).upper()
                ws.cell(y, x).fill = openpyxl.styles.PatternFill(fgColor=cor, fill_type='solid')

        wb.save(nome_arquivo + '.xlsx')

    except():
        print('An error occurred.')

# Chama a função main
if __name__ == '__main__':
    main()