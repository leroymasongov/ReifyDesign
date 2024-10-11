import json

# Define file path
input_file_path = r"C:\Users\leroy\Downloads\nwidd\CAD_interactions_report.json"

# Read the JSON file
with open(input_file_path, 'r') as json_file:
    visio_data = json.load(json_file)

# Iterate through each page and its components
for page in visio_data.get("pages", []):
    print(f"Page: {page['name']}")
    
    # Check nodes for "connect" in the name
    for node in page["components"].get("nodes", []):
        if node["name"] and "connect" in node["name"].lower():
            print(f"Node: {node}")
    
    # Check connections for "connect" in the name
    for connection in page["components"].get("connections", []):
        if connection["name"] and "connect" in connection["name"].lower():
            print(f"Connection: {connection}")