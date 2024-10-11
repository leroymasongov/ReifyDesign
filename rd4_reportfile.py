import win32com.client
from fileconfig import get_visio_file_path
from datetime import datetime

# Create a Visio application object
visio = win32com.client.Dispatch("Visio.Application")

# Get the Visio document path from the function
visio_file_path = get_visio_file_path()

# Open the Visio document
doc = visio.Documents.Open(visio_file_path)

# Get the current date and time for the file name
now = datetime.now()
short_date = now.strftime("%Y%m%d")
short_time = now.strftime("%H%M%S")
file_name = f"NWIDD-{short_date}-{short_time}.txt"
file_path = f"C:\\Users\\leroy\\Downloads\\{file_name}"

# Open the output file
with open(file_path, "w") as file:
    # Iterate through each page in the document
    for page in doc.Pages:
        file.write(f"Page: {page.Name}\n")
        
        # Iterate through each shape in the page
        for shape in page.Shapes:
            file.write(f"Shape ID: {shape.ID}\n")
            file.write(f"Shape Name: {shape.Name}\n")
            file.write(f"Shape Text: {shape.Text}\n")
            file.write(f"Shape Type: {shape.Type}\n")
            
            # Check if the shape is a connector
            if shape.OneD:
                try:
                    # Get the connected shapes
                    from_shape = None
                    to_shape = None
                    for connect in shape.Connects:
                        if connect.FromSheet:
                            from_shape = connect.FromSheet
                        if connect.ToSheet:
                            to_shape = connect.ToSheet
                    file.write(f"Connector from Shape ID: {from_shape.ID}, Name: {from_shape.Name}\n")
                    file.write(f"Connector to Shape ID: {to_shape.ID}, Name: {to_shape.Name}\n")
                except IndexError:
                    file.write("Error: Connector does not have the expected connections.\n")
            
            file.write("----\n")

# Close the document without saving
doc.Close(False)

# Quit the Visio application
visio.Quit()
