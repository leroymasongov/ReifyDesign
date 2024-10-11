import os
from PIL import Image, ImageFilter, ImageEnhance

# Define the input and output directories
input_dir = r'C:\Users\leroy\Downloads\nwidd\raw-images'
output_dir = r'C:\Users\leroy\Downloads\nwidd\mono-images'

# Ensure the output directory exists
os.makedirs(output_dir, exist_ok=True)

# Process each image in the input directory
for filename in os.listdir(input_dir):
    if filename.endswith(('.jpg', '.jpeg', '.png')):
        # Open the image
        img_path = os.path.join(input_dir, filename)
        img = Image.open(img_path)

        # Enhance the contrast to make the text clearer
        enhancer = ImageEnhance.Contrast(img)
        img_enhanced = enhancer.enhance(2)  # Increase the contrast
        # img_enhanced = img

        # Apply edge detection filter
        edges = img_enhanced.filter(ImageFilter.SHARPEN)

        # Convert the edges to pure black and white
        bw_edges = edges.convert()

        # Save the black and white edges image to the output directory
        output_path = os.path.join(output_dir, filename)
        bw_edges.save(output_path)


print("\n\n")
print("Conversion to mono shape outlines complete.")