import os
import cv2
import json
import sys
from collections import defaultdict

# Placeholder function to find objects and connections in the image
def find_objects_and_connections(image):
    objects = []
    connections = []
    return objects, connections

def process_images(input_dir):
    graph_data = defaultdict(list)
    
    for filename in os.listdir(input_dir):
        if filename.endswith(('.png', '.jpg', '.jpeg')):
            image_path = os.path.join(input_dir, filename)
            image = cv2.imread(image_path)
            objects, connections = find_objects_and_connections(image)
            graph_data[filename] = {'objects': objects, 'connections': connections}
    
    return graph_data

def save_graph_data(graph_data, output_file):
    with open(output_file, 'w') as f:
        json.dump(graph_data, f, indent=4)

if __name__ == "__main__":
    input_dir = r'C:\Users\leroy\Downloads\nwidd\mono-images'  # Define your input directory
    output_dir = r'C:\Users\leroy\Downloads\nwidd\json-graph'  # Define your output directory
    output_file = os.path.join(output_dir, 'graph_data.json')

    # Create the output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    graph_data = process_images(input_dir)
    save_graph_data(graph_data, output_file)

    # Restart the script
    os.execv(sys.executable, ['python'] + sys.argv)