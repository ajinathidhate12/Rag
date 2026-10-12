from abc import ABC, abstractmethod
from datetime import datetime
from pathlib import Path


class Document(ABC):

    def __init__(self, name, filesize, createdate, filetype):
        self.name = name
        self.filesize = filesize
        self.createdate = createdate
        self.filetype = filetype

    @abstractmethod
    def extract_text(self):
        pass

    @abstractmethod
    def clean_data(self):
        pass

    def __str__(self):
        return (
            f"Name: {self.name}, "
            f"Size: {self.filesize}, "
            f"Date: {self.createdate}, "
            f"Type: {self.filetype}"
        )

    
class xlsDocument(Document):

    def __init__(self, name, filesize, createdate, filetype, sheetcount):
        super().__init__(name, filesize, createdate, filetype)
        self.sheetcount = sheetcount
        self.filetype = "Excel"

    def extract_text(self):
        print("Extracting text from Excel")

    def clean_data(self):
        print("Cleaning Excel data")

class PDFDocument(Document):

    def __init__(self, name, filesize, createdate, filetype, pagecount):
        super().__init__(name, filesize, createdate, filetype)
        self.pagecount = pagecount
        self.filetype = "PDF"

    def extract_text(self):
        print("Extracting text from PDF")

    def clean_data(self):
        print("Cleaning PDF data")
    
class WordDocument(Document):

    def __init__(self, name, filesize, createdate, filetype, wordcount):
        super().__init__(name, filesize, createdate, filetype)
        self.wordcount = wordcount
        self.filetype = "Word"

    def extract_text(self):
        print("Extracting text from Word document")

    def clean_data(self):
        print("Cleaning Word document data")

class documentfactory:

    @staticmethod
    def create_document(doc_type, name, filesize, createdate, filetype, count):
        if doc_type == "PDF":
            return PDFDocument(name, filesize, createdate, filetype, count)
        elif doc_type == "Word":
            return WordDocument(name, filesize, createdate, filetype, count)
        elif doc_type == "Excel":
            return xlsDocument(name, filesize, createdate, filetype, count)
        else:
            raise ValueError("Invalid document type")


def demo_documents():
    doc1 = documentfactory.create_document("PDF", "sample.pdf", "2MB", "2023-01-01", "PDF", 10)
    doc2 = documentfactory.create_document("Word", "sample.docx", "1MB", "2023-01-01", "Word", 100)
    doc3 = documentfactory.create_document("Excel", "sample.xlsx", "1.5MB", "2023-01-01", "Excel", 5)

    print(doc1)
    print(doc2)
    print(doc3)


if __name__ == "__main__":
    demo_documents()
