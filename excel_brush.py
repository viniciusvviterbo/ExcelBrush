import sys, getopt
from PIL import Image

def print_instrucoes():
    print("excel_brush.py -s <size>")

def main():
    arquivo_imagem = ""
    tamanho_imagem = 0

    # Recebe as opções e argumentos passados como parâmetros
    try:
        opts, args = getopt.getopt(sys.argv,"hi:s:",["help", "image=", "size="])

    # Erro levantado quando opções inseridas não são reconhecidas
    except getopt.GetoptError:
        print_instrucoes()
        sys.exit(1)

    print(sys.argv)
    print(opts)
    print(args)
      
    for opt, arg in opts:
        if opt in ("-h", "--help"):
            print_instrucoes()
    
        elif opt in ("-i", "--image"):
            imagem_original = Image.open(arg)
        
        # Identifica a opção informada na lista de argumentos, posição 1
        elif opt in ("-s", "--size"):
            # Define o tamanho da resolução como o informado pelo usuário
            tamanho_imagem = int(arg)
    
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