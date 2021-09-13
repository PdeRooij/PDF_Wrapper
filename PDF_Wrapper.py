# PDF libraries
import pdfplumber
import pdfrw


class PDFWrapper:
    """
    Wrapper class with PDF functionalities.
    See https://akdux.com/python/2020/10/31/python-fill-pdf-files.html for excellent PDF form filling.
    """

    def __init__(self):
        """
        Initialize with PDF reader and writer. Prepare holders for reading and manipulation
        """
        # Initialize reader and writer
        self.reader = pdfplumber
        self.writer = pdfrw.PdfWriter

        # Prepare storing of pdf content
        self.pdf = None
        self.lib = None     # Remember what library is used for opening, to ensure correct funtionalities are used

        # Form keys for parsing
        self.ANNOT_KEY = '/Annots'
        self.ANNOT_FIELD_KEY = '/T'
        self.ANNOT_VAL_KEY = '/V'
        self.ANNOT_RECT_KEY = '/Rect'
        self.SUBTYPE_KEY = '/Subtype'
        self.WIDGET_SUBTYPE_KEY = '/Widget'

    def read_pdf(self, path, lib='pdfplumber'):
        """
        Open a connection to specified PDF for later queries.

        Args:
            path (str): Location of the PDF file to be read.
            lib (str): (Optional) Library to use for opening. Defaults to pdfplumber.
        """
        self.lib = lib
        if lib == 'pdfrw':
            self.pdf = pdfrw.PdfReader(path)
        else:
            self.pdf = pdfplumber.open(path)

    def extract_fields_values(self, filled_path):
        """
        Puts names (key) and values (value) of all form fields in a pdf into a dictionary.
        Note that fields without a value have None as dictionary value.

        Args:
            filled_path (str): Location of the PDF to extract fields and values from.

        Returns:
            dict: Dictionary of fields and values, where field names are keys.
        """
        # Prepare dictionary
        fields = {}

        # Read PDF and extract form field names and values
        filled = pdfrw.PdfReader(filled_path)
        for page in filled.pages:
            annotations = page[self.ANNOT_KEY]
            for annotation in annotations:
                if annotation[self.SUBTYPE_KEY] == self.WIDGET_SUBTYPE_KEY:
                    if annotation[self.ANNOT_FIELD_KEY]:
                        if annotation[self.ANNOT_VAL_KEY]:
                            fields[annotation[self.ANNOT_FIELD_KEY][1:-1]] = annotation[self.ANNOT_VAL_KEY]
                        else:
                            fields[annotation[self.ANNOT_FIELD_KEY][1:-1]] = None

        # Return extracted field names and values
        return fields

    def fill_pdf(self, template_path, output_path, field_value_dict):
        template_pdf = pdfrw.PdfReader(template_path)
        for page in template_pdf.pages:
            annotations = page[self.ANNOT_KEY]
            for annotation in annotations:
                if annotation[self.SUBTYPE_KEY] == self.WIDGET_SUBTYPE_KEY:
                    if annotation[self.ANNOT_FIELD_KEY]:
                        key = annotation[self.ANNOT_FIELD_KEY][1:-1]
                        if key in field_value_dict.keys():
                            if type(field_value_dict[key]) == bool:
                                if field_value_dict[key] == True:
                                    annotation.update(pdfrw.PdfDict(
                                        AS=pdfrw.PdfName('Yes')))
                            else:
                                annotation.update(
                                    pdfrw.PdfDict(V='{}'.format(field_value_dict[key]))
                                )
                                annotation.update(pdfrw.PdfDict(AP=''))
        template_pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))  # NEW
        pdfrw.PdfWriter().write(output_path, template_pdf)

    def close_pdf(self):
        """
        Close connection to PDF that was opened earlier.
        """
        # Only if pdfplumber is used, does the pdf have to be closed
        if self.lib == 'pdfplumber':
            self.pdf.close()
