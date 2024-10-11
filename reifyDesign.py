import win32com.client

# Create a Visio application object
visio = win32com.client.Dispatch("Visio.Application")

# Open the Visio document
doc =''
# visio.Documents.Open("C:\\Users\\leroy\Downloads\\CAD_interactions.vsdx")

# Iterate through each page in the document
for page in doc.Pages:
    print(f"Page: {page.Name}")
    
    # Iterate through each shape in the page
    for shape in page.Shapes:
        print(f"Shape ID: {shape.ID}")
        print(f"Shape Name: {shape.Name}")
        print(f"Shape Text: {shape.Text}")
        print(f"Shape Type: {shape.Type}")
        print("----")

# Close the document without saving
doc.Close()

# Quit the Visio application
visio.Quit()