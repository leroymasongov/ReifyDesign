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
        
        page_data["shapes"].append(shape_data)
        
        if shape_type == "Connector":
            page_data["connections"].append(shape_data)
        else:
            page_data["nodes"].append(shape_data)
    
    return page_data

def extract_document_data(document):
    document_data = {
        "pages": []
    }
    
    for page in document.pages:
        page_data = extract_page_data(page)
        document_data["pages"].append(page_data)
    
    return document_data

# Read the Visio file
with VisioFile(input_file_path) as visio:
    document_data = extract_document_data(visio)

# Write the extracted data to a JSON file
with open(output_file_path, 'w') as f:
    json.dump(document_data, f, indent=4)