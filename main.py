import docx
from docx.enum.text import (WD_PARAGRAPH_ALIGNMENT,
                            WD_BREAK_TYPE,
                            WD_COLOR_INDEX,
                            WD_LINE_SPACING,
                            WD_TAB_ALIGNMENT,
                            WD_TAB_LEADER,
                            WD_UNDERLINE)
from docx.enum.style import (WD_BUILTIN_STYLE,
                             WD_STYLE_TYPE)
from docx.enum.dml import (MSO_COLOR_TYPE,
                           MSO_THEME_COLOR_INDEX)
from docx.shared import Pt, Cm, Mm, RGBColor
from socket import gethostname
from os import path
from sys import argv
import json
import re


class Report:
    """ Provides methods for working with the generated report in accordance with the state standard """
    def __init__(self, filename: str, styles_path="styles.json"):
        self.styles_path = path.abspath(styles_path)
        self.filename = filename
        self.check_filename()

        try:
            document = docx.Document()
            document.save(filename)
            self.document = docx.Document(filename)
        except PermissionError:
            exit("To generate a report, close the file that you specified for recording")

    def __call__(self):
        self.set_styles()
        self.parse_content()

    def check_filename(self):
        """ Checks if the file path and file name are correct """
        if re.match(r"^(/+|\\+)$", self.filename) is not None:
            exit("Wrong file path")

        dlm: str = '/' if self.filename.find("/") != -1 else "\\"
        if not path.exists(dlm.join(self.filename.split(dlm)[:-1:])):
            exit("File path is specified incorrectly")

        if len(self.filename.split(".")) < 2 or self.filename.split(".")[-1] != "docx":
            exit("Incorrect file format")

        """if path.exists(filename):
            if input("File was found at the specified path. Do you really want to overwrite it? (Y/N)\n") not in ["Y", "y"]:
                exit()"""

    def set_styles(self):
        """ Adds all the necessary styles from .json file"""
        if not path.exists(self.styles_path):
            exit("The file containing the styles was not found")

        with open(self.styles_path, "r", encoding="utf-8") as file:
            styles_from_file = json.load(file)

        styles = self.document.styles
        for style in styles:
            styles[style.name].delete()

        for (name, values) in styles_from_file["styles"].items():
            current_style = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            current_style.font.name = values["font"]["name"]
            current_style.font.size = Pt(values["font"]["size"])

            current_style.font.all_caps = values["font"]["all_caps"]
            current_style.font.bold = values["font"]["bold"]
            current_style.font.italic = values["font"]["italic"]
            current_style.font.underline = values["font"]["underline"]

            current_style.font.color.ColorFormat = MSO_COLOR_TYPE.RGB
            current_style.font.color.rgb = RGBColor(*values["font"]["color"])

            current_style.font.math = values["font"]["math"]
            current_style.font.no_proof = values["font"]["no_proof"]

    def parse_content(self):
        """ Parses files in the directory and composes the document from them """
        self.document.add_heading("Heading 1", 1)
        """self.document.add_paragraph("Heading 2", style="Heading 2")
        self.document.add_paragraph("Heading 3", style="Heading 3")
        self.document.add_paragraph("NORMal Text herrrre", style="Normal")"""
        self.document.core_properties.revision += 1
        self.document.save(self.filename)


# python main.py documents/testReport.docx
def main():
    try:
        filename = argv[1]
    except IndexError:
        filename = input('Enter file name: ')

    report = Report(path.abspath(filename))
    report()


if __name__ == '__main__':
    main()
