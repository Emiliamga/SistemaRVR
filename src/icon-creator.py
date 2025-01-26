# create_icon.py
from PIL import Image, ImageDraw, ImageFont

def create_icon():
    # Criar uma imagem 256x256 com fundo transparente
    size = (256, 256)
    image = Image.new('RGBA', size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(image)
    
    # Desenhar um círculo como fundo
    circle_color = (65, 105, 225)  # Azul royal
    center = (128, 128)
    radius = 120
    draw.ellipse([center[0]-radius, center[1]-radius, 
                  center[0]+radius, center[1]+radius], 
                 fill=circle_color)
    
    # Adicionar texto "RD" (Relatório de Despesas)
    try:
        font = ImageFont.truetype("Arial.ttf", 100)
    except:
        font = ImageFont.load_default()
    
    text = "RD"
    text_color = (255, 255, 255)  # Branco
    
    # Centralizar o texto
    text_bbox = draw.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]
    text_position = (128 - text_width//2, 128 - text_height//2)
    
    draw.text(text_position, text, fill=text_color, font=font)
    
    # Salvar em diferentes formatos
    image.save("icone.png")
    image.save("icone.ico", format="ICO", sizes=[(256, 256)])

if __name__ == "__main__":
    create_icon()
