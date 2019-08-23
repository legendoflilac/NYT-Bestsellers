# NYT-Bestsellers Reworked
A Python script for creating a Word document with data from three Bestseller lists--now using the New York Times API, found at https://developer.nytimes.com/.

This update was necessary because of code changes on the NYT website. The old script can be used if updated with the correct selectors, but ultimately this one is faster and will likely have fewer issues running in the future.


# Ideas for future improvements
- Rename output file so it adds the date to the file name
- Add a header with the date to the document
- Search the online catalog to see if the titles exist in the library's collection, and if so, populate the leftmost column with the call numbers

# Bugs
- The docx file displays best on newer versions of MS Word. Tried opening in Word 2010 and the header of the second table was over the header of the first table. The file looks fine on newer versions.


# Old Description
A Python script for populating a table within a Word document with data from three of the New York Times Bestseller lists.

Using requests, BeautifulSoup4, and python-docx, I constructed a script that fills in the title, author, publisher, description, and "freshness" (how many weeks a book has been on the list) of a listed book into a template. 
