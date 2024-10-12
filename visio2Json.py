import json
from vsdx import VisioFile

# Define file paths
input_file_path = r"C:\Users\leroy\Downloads\nwidd\CAD_interactions.vsdx"
output_file_path = r"C:\Users\leroy\Downloads\nwidd\CAD_interactions_report.json"

# Function to extract shapes and connections from a Visio page
def extract_page_data(page):
    page_data = {
        "page_name": page.name,
        "shapes": [],
        "nodes": [],
        "connections": []
    }
    
    # Extract shapes and categorize them into nodes and connections
    for shape in page.all_shapes:
        # Determine the shape type
        width = shape.width
        height = shape.height
        if shape.shape_name and "connect" in shape.shape_name.lower():
            shape_type = "Connector"
        elif abs(width - height) < 1e-2:  # Tolerance for floating-point comparison
            shape_type = "Circular"
        else:
            shape_type = "Rectangular"
        
        shape_data = {
            "id": shape.ID,
            "name": shape.shape_name,
            "text": shape.text,
            "type": shape.shape_type,
            "shape_type": shape_type,
            "page_name": page.name
        }
        
        if shape.shape_name and "connector" in shape.shape_name.lower():
            # Find the from and to shapes
            from_shape_id = shape.connects[0].from_id if shape.connects else None
            to_shape_id = shape.connects[0].to_id if shape.connects else None
            
            # Split the text by commas if it contains commas
            text_values = shape.text.split(',') if ',' in shape.text else [shape.text]
            
            for text_value in text_values:
                connector_data = shape_data.copy()
                connector_data["text"] = text_value.strip()
                connector_data["from_id"] = from_shape_id if from_shape_id else None
                connector_data["to_id"] = to_shape_id if to_shape_id else None
                
                page_data["connections"].append(connector_data)
        else:
            page_data["nodes"].append(shape_data)
        
        page_data["shapes"].append(shape_data)
    
    return page_data

# Read the Visio file
with VisioFile(input_file_path) as visio:
    data = []
    for page in visio.pages:
        page_data = extract_page_data(page)
        data.append(page_data)

# Write the extracted data to a JSON file
with open(output_file_path, 'w') as f:
    json.dump(data, f, indent=4)