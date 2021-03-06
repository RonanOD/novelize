# Novelize
Python script to transform docx files created by the [novelize website](https://getnovelize.com/).
I used novelize to write my novel [Chief O'Neill](http://chiefoneill.com/the-book). 

# Details
 * Script to add a front image to a novelize document. Front image file must be `images/frontImage.jpg`.
 * Also add any images with "``__*imageName*__``" markup in a paragraph in Novelize.
 * All images go in the images folder
 * Call script with doc name as argument: `python novelize.py my-manuscript.docx`
 * Output is `final-my-manuscript.docx` with transformations included.
 * Check header of novelize.py for further possible transformations.

# Useful links:
 * python-docx - https://python-docx.readthedocs.io/en/latest/index.html
 * replacing text - https://www.quora.com/How-can-I-find-and-replace-text-in-a-Word-document-using-Python
