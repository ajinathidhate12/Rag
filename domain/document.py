from abc import ABC, abstractmethod

class Document():

    """A class representing a document with its name, filesize, creation date and file type.  """
    """base document class and the derived class will be in the test2.py file"""

    def __init__(self, name, filesize, createdate,filetype):
        self.name = name
        self.filesize = filesize
        self.createdate = createdate
        self.filetype = filetype

