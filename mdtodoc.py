import markdown
import argparse
import sys
from spire.doc import *
from spire.doc.common import *

### take arguments, get filename
parser = argparse.ArgumentParser(
    epilog="\tExample: \r\npython3 "
    + sys.argv[0]
    + " -f Random_Markdown_File.md -o output_file_in_doc_format.doc"
)
parser._optionals.title = "OPTIONS"
parser.add_argument(
    "-f",
    "--file",
    help="Specify input file, must be markdown",
    required=True,
)
parser.add_argument(
    "-o",
    "--outputfile",
    default="md_output.doc",
    help="optional ouput name to prepend to scan outputs.",
)

args = parser.parse_args()
if args.file is not None:
    inputfile = str(args.file)
else:
    print("not input file provided.")


output_file = args.outputfile


def markdown_to_docx(input_file, output):
    # Convert Markdown to HTML
    html = markdown.markdown(input_file)

    # Create a new Document
    doc = Document()
    # doc.LoadFromFile(html, FileFormat.Html, XHTMLValidationType.none)
    sec = doc.AddSection()
    paragraph = sec.AddParagraph()
    paragraph.AppendHTML(html)
    doc.SaveToFile(output, FileFormat.Docx2016)
    doc.Close()


# Read Markdown content from a file
with open(inputfile, "r", encoding="utf-8") as file:
    input_data = file.read()

# Convert and save to .docx
markdown_to_docx(input_data, output_file)
