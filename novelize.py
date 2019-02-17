from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_BREAK
import os, re, sys
###
# Script to add a front image to a novelize document. Front image file must be
# images/frontImage.jpg.
# Also add any images with "__*imageName*__" markup in a paragraph in novelize.
# All images go in the images folder
# Call script with doc name as argument: python novelize.py my-manuscript.docx
# Output is final-my-manuscript.docx with transformations included.
#
# Useful links:
# python-docx - https://python-docx.readthedocs.io/en/latest/index.html
# replacing text - https://www.quora.com/How-can-I-find-and-replace-text-in-a-Word-document-using-Python
###

# Variables and Constants
docName = ''
finalDocName = ''
frontImage = 'images/frontImage.jpg'
regex = re.compile(r"__\*(.*)\*__")
imageWidth = 6.5

# Start up
if len(sys.argv) < 2:
    print("Run script with document name as argument: novelize.py my-manuscript.docx")
else:
    docName = sys.argv[1]
    finalDocName = 'final-' + docName

# Remove final doc if necessary
if os.path.exists(finalDocName):
    os.remove(finalDocName)
# Open the document
document = Document(docName)
# Add front image
para = document.paragraphs[0].insert_paragraph_before()
run = para.add_run()
run.add_picture(frontImage, width=Inches(imageWidth))
run.add_break(WD_BREAK.PAGE)

# Look for __*imageName*__ in a paragraph on its own 
# and replace with imageName in same folder:
for p in document.paragraphs:
    match = regex.search(p.text)
    if match:
        imageFile = "images/" + match.group(1)
        for run in p.runs: # delete the text
            run.text = ""
        lastRun = p.add_run()
        lastRun.add_picture(imageFile, width=Inches(imageWidth))

# Finished. Save the final document
document.save(finalDocName)