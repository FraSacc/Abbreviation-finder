#!/usr/bin/env python
# coding: utf-8

# In[9]:


'''A script that finds all abbreviations in a text, composed of uppercase letters,
hyphens, slashes and Greek letters, or a combination of these.
It is still advisable to check the final list for omissions or misidentified abbreviations '''

import docx
import re
import csv

def extract_abbreviations(filename):
    # Open the Word document
    doc = docx.Document(filename)

    # Initialize a list to store the abbreviations
    abbreviations = []

    # Compile a regular expression to match uppercase abbreviations with combinations of hyphens, slashes, and Greek letters
    # (e.g. "ABC", "AB-C", "AB/C", "ABCΓ", "AB-CΓ", "A-BC", "AΓ-BC", "C/EBPα", "ABΓ-C", "AΓ-BC")
    abbrev_pattern = re.compile(r'[A-Z]{2,}(?:[-/][A-Z]{2,}|[Α-ΩΆ-Ώ])*(?:[-/][A-Z]{2,}|[Α-ΩΆ-Ώ])*')

    # Iterate through the paragraphs in the document
    for paragraph in doc.paragraphs:
        # Split the paragraph into words
        words = paragraph.text.split()
       
        # Iterate through the words in the paragraph
        for word in words:
            # Check if the word matches the abbreviation pattern
            if abbrev_pattern.match(word):
                # Add the word to the list of abbreviations
                if word not in abbreviations and int(len(word))<=7 and '.' not in word and ',' not in word and '(' not in word and ')' not in word and ':' not in word:
                    abbreviations.append(word)

    # Return the list of abbreviations
    return abbreviations

# Example usage:
abbreviations = extract_abbreviations("Document.docx")
print("My list of abbreviations:",abbreviations)

with open('Abbreviations', 'w',encoding='utf-8') as f:
      
    # using csv.writer method from CSV package
    write = csv.writer(f)
    write.writerows([abbreviation]for abbreviation in abbreviations)


# In[ ]:




