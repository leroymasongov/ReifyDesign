import win32com.client

# Create a Visio application object
visio = win32com.client.Dispatch("Visio.Application")

# Open the Visio document
doc = visio.Documents.Open("C:\\Users\\leroy\Downloads\\CAD_interactions.vsdx")

# Iterate through each page in the document
for page in doc.Pages:
    print(f"Page: {page.Name}")
    
    # Iterate through each shape in the page
    for shape in page.Shapes:
        print(f"Shape ID: {shape.ID}")
        print(f"Shape Name: {shape.Name}")
        print(f"Shape Text: {shape.Text}")
        print(f"Shape Type: {shape.Type}")
        
        # Check if the shape is a connector
        if shape.OneD:
            # Get the connected shapes
            from_shape = shape.Connects.Item(1).ToSheet
            to_shape = shape.Connects.Item(2).ToSheet
            print(f"Connector from Shape ID: {from_shape.ID}, Name: {from_shape.Name}")
            print(f"Connector to Shape ID: {to_shape.ID}, Name: {to_shape.Name}")
        
        print("----")

# Close the document without saving
doc.Close(False)

# Quit the Visio application
visio.Quit()