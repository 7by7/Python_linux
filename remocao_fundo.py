from rembg import remove # type: ignore
from PIL import Image

url = Image.open('Botom Metodo MTR.png')
output = remove(url)
output.save('logo_MTR.png')