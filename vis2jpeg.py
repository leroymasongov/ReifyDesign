import os
import psutil
from fileconfig import get_visio_file_path
import comtypes.client
from PIL import Image

def kill_visio_processes():
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and 'visio' in proc.info['name'].lower():
            proc.terminate()
            proc.wait()

def convert_visio_to_jpeg():
    kill_visio_processes()
    
    visio_file = get_visio_file_path()
    visio_app = comtypes.client.CreateObject('Visio.Application')
    visio_app.Visible = False

    visio_doc = visio_app.Documents.Open(visio_file)
    output_folder = r'C:\Users\leroy\Downloads\nwidd\raw-images'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for i, page in enumerate(visio_doc.Pages):
        output_file = os.path.join(output_folder, f'diagram_{i + 1}.png')
        page.Export(output_file)

        # Open the exported image and resize it
        with Image.open(output_file) as img:
            new_size = (img.width * 4, img.height * 4)
            resized_img = img.resize(new_size, Image.Resampling.LANCZOS)
            resized_img.save(output_file)

    visio_doc.Close()
    visio_app.Quit()

if __name__ == "__main__":
    convert_visio_to_jpeg()