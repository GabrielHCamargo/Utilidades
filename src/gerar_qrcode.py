import qrcode
import os
 
# Solicita ao usuário que insira a URL para gerar o QR code
url = input("Digite a URL para gerar o QR code: ")

# Cria um objeto QRCode com configurações específicas
qr = qrcode.QRCode(version=1, box_size=10, border=5)
qr.add_data(url)
qr.make(fit=True)

# Gera a imagem do QR code
img = qr.make_image(fill_color="black", back_color="white")

# Define o diretório temporário para salvar o QR code
temp_dir = ".temp"
if not os.path.exists(temp_dir):
    os.makedirs(temp_dir)

# Gera o nome do arquivo baseado na URL fornecida
file_name = url.replace("https://", "").split("/")[0] + ".png"
file_path = os.path.join(temp_dir, file_name)

# Salva a imagem do QR code no arquivo
img.save(file_path)

# Exibe o caminho completo onde o QR code foi salvo
print(f"O QR code foi salvo em: {os.path.abspath(file_path)}")
