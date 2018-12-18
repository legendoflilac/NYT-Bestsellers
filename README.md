# NYT-Bestsellers
A Python script for populating a table within a Word document with data from three of the New York Times Bestseller lists.

Using requests, BeautifulSoup4, and python-docx, I constructed a script that fills in the title, author, publisher, description, and "freshness" (how many weeks a book has been on the list) of a listed book into a template. 

# Ideas for future improvements
- Rename output file so it adds the date to the file name
- Add a header with the date to the document
- Search the online catalog to see if the titles exist in the library's collection, and if so, populate the leftmost column with the call numbers
