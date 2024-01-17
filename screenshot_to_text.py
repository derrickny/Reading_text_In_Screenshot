import cv2
import pytesseract
import pandas as pd
import os
import re
import numpy as np
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT

# Path to the Word document
doc_path = '/Users/nyagaderrick/Downloads/Oracle user accounts.docx'

# Initialize an empty DataFrame to store all the text
all_text = pd.DataFrame()

# Load the Word document
doc = Document(doc_path)


# Loop through all the images in the Word document
for rel in doc.part.rels.values():
    if "image" in rel.reltype:
        image_data = rel._target._blob  # This is the raw image datahis is the raw image data

        # Convert the image data to an OpenCV image
        nparr = np.frombuffer(image_data, np.uint8)
        image = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

        # Convert the image to grayscale
        gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        # Use pytesseract to extract text
        pytesseract.pytesseract.tesseract_cmd = r'/opt/homebrew/bin/tesseract'  # Update this path to your Tesseract executable
        text = pytesseract.image_to_string(gray_image)

        # Clean and process the text as necessary
        lines = text.split('\n')
        data = []

        for line in lines:
            # Define regular expressions for each piece of information
            name_regex = r"idcsdev(\w*)"
            status_regex = r"status\s*(yes)"
            email_regex = ""  # Email is blank for all
            federated_regex = r"federated\s*(yes)"
            created_regex = r"created\s*((?:mon|tue|wed|thu|fri|sat|sun)\s*\w*)"

            # Use re.search() to find matches in the text
            name_match = re.search(name_regex, line, re.IGNORECASE)
            status_match = re.search(status_regex, line, re.IGNORECASE)
            federated_match = re.search(federated_regex, line, re.IGNORECASE)
            created_match = re.search(created_regex, line, re.IGNORECASE)

            # Extract the matched text and add it to the data list
            if name_match and status_match and federated_match and created_match:
                name = name_match.group(1)
                status = status_match.group(1)
                email = ""  # Email is blank for all
                federated = federated_match.group(1)
                created = created_match.group(1)
                data.append([name, status, email, federated, created])

        # Store the text in a pandas DataFrame
        df = pd.DataFrame(data, columns=['Name', 'Status', 'Email', 'Federated', 'Created'])

        # Append the text to the all_text DataFrame
        all_text = pd.concat([all_text, df])

# Write the all_text DataFrame to an Excel file
all_text.to_excel('output.xlsx', index=False)