# import mammoth
# import pdfkit
# from docx2pdf import convert
import subprocess
from pathlib import Path

# f = open("Sails Software JD and Interview Questions.docx", "rb")
# b = open("Sails Software JD and Interview Questions.docx.html", "wb")
# document = mammoth.convert_to_html(f)
# b.write(document.value.encode("utf8"))
# f.close()
# b.close()


# convert("Sails Software JD and Interview Questions.docx")


def convert_docx_to_pdf(input_path, output_folder=None):
    input_path = Path(input_path).resolve()  # Ensure absolute path

    # Check if the input file exists
    if not input_path.is_file():
        print(f"❌ Error: The file '{input_path}' does not exist.")
        return

    # Default output folder to current working directory if not provided
    if output_folder is None:
        output_folder = Path.cwd()
    else:
        output_folder = Path(output_folder).resolve()  # Ensure absolute path

    # Check if the output folder exists
    if not output_folder.is_dir():
        print(f"❌ Error: The output folder '{output_folder}' does not exist.")
        return

    # Execute the conversion command
    command = [
        "soffice",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_folder),  # Convert Path object to string
        str(input_path),
    ]

    subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    print(f"✅ PDF saved in: {output_folder}")


# Example usage with absolute path for output folder
convert_docx_to_pdf(
    "pythonConverter/input.docx",
    "/Users/YasinKhan/coding/convert-docx-to-pdf/pythonConverter",
)
