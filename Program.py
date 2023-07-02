import os
import pdfplumber
import pandas as pd

# Path to the folder containing PDF files
folder_path = "C:/Users/Admin/Desktop/PDF task/Stress Test 2"

# Initialize an empty list to store the word count for each file
word_counts = []

# Loop through each file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):
        # Open the PDF file
        with pdfplumber.open(os.path.join(folder_path, filename)) as pdf:
            # Get the total number of pages in the PDF file
            num_pages = len(pdf.pages)

            # Loop through each page and count the number of words
            word_count = 0
            for i, page in enumerate(pdf.pages):
                # Skip the first page for specific files
                if filename in ["Santander Holdings USA 2015.pdf", "Santander Holdings USA 2016.pdf", 
                                 "Santander Holdings USA 2017.pdf", "Santander Holdings USA 2018.pdf", 
                                 "Santander Holdings USA July 2015.pdf", "Santander Holdings USA March 2014.pdf", 
                                 "Santander Holdings USA Mid-Cycle 2016.pdf", "Santander Holdings USA Mid-Cycle 2017.pdf", 
                                 "Santander Holdings USA Mid-Cycle 2018.pdf", "Santander USA 2014.pdf", 
                                 "SVB Financial 2017.pdf"] and i == 0:
                    continue
                text = page.extract_text()
                word_count += len(text.split())

            # Store the word count for the file in the list
            word_counts.append((filename, word_count))

# Convert the list of tuples to a pandas DataFrame
df = pd.DataFrame(word_counts, columns=["Filename", "Word Count"])

# Save the DataFrame to an Excel file
df.to_excel("Stress Test 2 word_counts.xlsx", index=False)