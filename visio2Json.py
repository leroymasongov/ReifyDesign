import json
from vsdx import VisioFile

# Define file paths
input_file_path = r"C:\Users\leroy\Downloads\nwidd\CAD_interactions.vsdx"
output_file_path = r"C:\Users\leroy\Downloads\nwidd\CAD_interactions_report.json"

# Function to extract shapes and connections from a Visio page
def extract_page_data(page):
    page_data = {
        "shapes": [],
        "nodes": [],
        "connections": []
    }
    
    # Extract shapes and categorize them into nodes and connections
    for shape in page.all_shapes:
        shape_data = {
            "id": shape.ID,
            "name": shape.shape_name,
            "text": shape.text,
            "type": shape.shape_type
        }
        
        if shape.shape_name and "connector" in shape.shape_name.lower():
            # Find the from and to shapes
            from_shape_id = shape.connects[0].from_id if shape.connects else None
            to_shape_id = shape.connects[0].to_id if shape.connects else None
            
            # Add from and to shape IDs to the shape data
            shape_data["from_id"] = from_shape_id if from_shape_id else None
            shape_data["to_id"] = to_shape_id if to_shape_id else None
            
            page_data["connections"].append(shape_data)
        else:
            page_data["nodes"].append(shape_data)
        
        page_data["shapes"].append(shape_data)
    
    return page_data

# Read the Visio file
with VisioFile(input_file_path) as visio:
    visio_data = {
        "pages": []
    }
    
    # Iterate through each page
    for page in visio.pages:
        page_data = {
            "name": page.name,
            "components": extract_page_data(page)
        }
        visio_data["pages"].append(page_data)

# Write the data to a JSON file
with open(output_file_path, 'w') as json_file:
    json.dump(visio_data, json_file, indent=4)

print(f"\n\n\nJSON report generated at {output_file_path}\n")