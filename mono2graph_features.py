import os
import cv2
import json
import sys
from collections import defaultdict

# Load pre-trained model and configuration for object detection
net = cv2.dnn.readNetFromCaffe('deploy.prototxt', 'res10_300x300_ssd_iter_140000.caffemodel')

def find_objects_and_connections(image):
    (h, w) = image.shape[:2]
    blob = cv2.dnn.blobFromImage(cv2.resize(image, (300, 300)), 1.0,
                                 (300, 300), (104.0, 177.0, 123.0))
    net.setInput(blob)
    detections = net.forward()

    objects = []
    connections = []  # Placeholder for connections logic

    for i in range(detections.shape[2]):
        confidence = detections[0, 0, i, 2]
        if confidence > 0.5:
            box = detections[0, 0, i, 3:7] * np.array([w, h, w, h])
            (startX, startY, endX, endY) = box.astype("int")
            objects.append({
                'confidence': float(confidence),
                'box': [int(startX), int(startY), int(endX), int(endY)]
            })

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