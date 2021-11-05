from PIL import Image, ImageFont, ImageDraw 
import os

# 25.4 × 19.05 cm


def add_name_to_image(name):
    """
    Add a name to a PDF file.
    """
    filepath = os.path.abspath("assets/Both_tracks_name.pdf")
    image = Image.open(filepath)
    caveat_font = ImageFont.truetype('Caveat/caveat.ttf', 50)
    title_text = name
    image_editable = ImageDraw.Draw(image)
    image_editable.text((60,285), title_text, (0, 0, 0), font=caveat_font)
    image.save(f"output/{name}_certificate.pdf")


add_name_to_image("Akhil Bhalerao")