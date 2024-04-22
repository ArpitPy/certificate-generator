import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
from matplotlib.backends.backend_pdf import PdfPages

# Function to calculate the width of the text
def calculate_text_width(text):
    # Estimate width of each character as 20px and space between each character as 2px
    character_width = 20
    space_between_characters = 2
    # Estimate space between two words as 10px
    space_between_words = 10

    # Calculate total width of the text
    total_width = (len(text) * character_width) + ((len(text) - 1) * space_between_characters)

    return total_width

data = pd.read_excel(r'PATH TO EXCEL SHEET')

# Create a PDF pages object
pdf_pages = PdfPages("all_certificates.pdf")

for index, row in data.iterrows():
    name = row["Name"]
    year = row["Year"]
    dept = row["Department"]
    sport = row["Sport"]
    
    im = Image.open(r'PATH TO BLANK CERTIFICATE')
    d = ImageDraw.Draw(im)
    
    font = ImageFont.truetype("arial.ttf", 27)
    
    name_location = (614, 568)
    year_location = (1164, 568)
    dept_location = (536, 635)
    sport_location = (1070, 633)
    
    text_color = (0, 0, 0)
    
    # Calculate the width of the text
    name_width = calculate_text_width(name)
    year_width = calculate_text_width(year)
    dept_width = calculate_text_width(dept)
    sport_width = calculate_text_width(sport)
    
    # Calculate the starting position to center the text
    name_starting_x = name_location[0] + ((name_width - calculate_text_width(' ')) / 2)
    year_starting_x = year_location[0] + ((year_width - calculate_text_width(' ')) / 2)
    dept_starting_x = dept_location[0] + ((dept_width - calculate_text_width(' ')) / 2)
    sport_starting_x = sport_location[0] + ((sport_width - calculate_text_width(' ')) / 2)
    
    d.text((name_starting_x, name_location[1]), name, fill=text_color, font=font)
    d.text((year_starting_x, year_location[1]), year, fill=text_color, font=font)
    d.text((dept_starting_x, dept_location[1]), dept, fill=text_color, font=font)
    d.text((sport_starting_x, sport_location[1]), sport, fill=text_color, font=font)
    
    # Save the image as BytesIO object
    buffer = BytesIO()
    im.save(buffer, format="JPEG")
    buffer.seek(0)
    
    # Create a new PIL Image object from the BytesIO buffer
    image = Image.open(buffer)
    
    # Add the current certificate page to the PDF
    pdf_pages.savefig(image)

# Close the PDF pages object
pdf_pages.close()
