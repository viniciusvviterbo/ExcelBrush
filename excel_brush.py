from PIL import Image
import argparse

# Configuracao dos argumentos
parser = argparse.ArgumentParser(description = 'Um programa para desenhar fotos em planilhas.')
parser.add_argument('-i', action = 'store', dest = 'image',
                    default = '', required = True,
                    help = 'A imagem a ser processada pelo programa.')
parser.add_argument('-s', action = 'store', dest = 'outputImageSize', required = True,
                    help = 'O tamanho da imagem a ser retornada')


def print_instrucoes():
    print("excel_brush.py -s <size>")

def main():
    arquivo_imagem = ""
    tamanho_imagem = 0

    # Recebe os argumentos, se as variaveis nao forem passadas, retorna -h
    arguments = parser.parse_args()
    imagem_original = Image.open(arguments.image)
    tamanho_imagem = int(arguments.outputImageSize)

        
    try:
        # Cria uma imagem reduzida que é nada mais que a imagem_original redimensionada com a resolução informada
        imagem_reduzida = imagem_original.resize((tamanho_imagem, tamanho_imagem), Image.BILINEAR)
        # Cria a imagem final redimensionando a imagem_reduzida ao tamanho da original
        imagem_pixelada = imagem_reduzida.resize(imagem_original.size, Image.NEAREST)
        # Define o nome do arquivo recém-criado
        nome_arquivo = "andre_" + str(tamanho_imagem) + "x" + str(tamanho_imagem) + ".png"
        # Salva a imagem criada
        imagem_pixelada.save(nome_arquivo, dpi=(200,200))

    except():
        print('An error occurred.')

# Chama a função main
if __name__ == "__main__":
    main()