from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.pagesizes import letter
from PIL import Image
import os, datetime
import pandas as pd

from gen_chart import save_chart

df = pd.read_excel('df.xlsx', header=0)
print(df)


# Paths and variables
template_path = "Merk_Diverse_Suppliers.pdf"
output_dir = "Generated PDFs"
term = 'Fall 2077'

# Ensure output directory exists
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    name = row['Name']  # Using 'Name' column as the identifier
    print(index)
    save_chart(row, df)
    # break

    print(f"Generating PDF for {name}")

    # # Create the overlay for each employee
    # o = canvas.Canvas("data_overlay.pdf", pagesize=letter)
    # o.setFont("Helvetica", 11)
    # o.setFillColor("white")
    # o.drawString(62, 564, name + ', ' + term)

    # # Read the template PDF
    # template = PdfReader(template_path)

    # # Loop through pages of the template and overlay charts
    # for pagenum in range(len(template.pages) - 1):
    #     o.showPage()


    #     # Place charts on page 2
    #     if pagenum == 2:
    #         try:
    #             chart1 = ImageReader(f"{output_dir}/MOTAI{index + 1}.png")
    #             o.drawImage(chart1, 53, 447, 250, 150, mask="auto")
                
    #             chart2 = ImageReader(f"{output_dir}/MOTSN{index + 1}.png")
    #             o.drawImage(chart2, 308, 447, 250, 150, mask="auto")
                
    #             chart3 = ImageReader(f"{output_dir}/MOTNC{index + 1}.png")
    #             o.drawImage(chart3, 53, 170, 250, 150, mask="auto")
                
    #             chart4 = ImageReader(f"{output_dir}/Leff{index + 1}.png")
    #             o.drawImage(chart4, 308, 170, 250, 150, mask="auto")
    #         except Exception as e:
    #             print(f"Error placing charts for {name}: {e}")

    

    #     o.setFont("Helvetica", 10)
    #     if pagenum != 24:
    #         o.drawString(535 - o.stringWidth(name.split('.')[0]), 38.3, name.split('.')[0])

    # o.save()

    # # Read the overlay and merge with the template
    # overlay = PdfReader("data_overlay.pdf")
    # writer = PdfWriter()

    # for pagenum in range(len(template.pages)):
    #     target_page = template.pages[pagenum]

    #     try:
    #         overlay_page = overlay.pages[pagenum]
    #         target_page.merge_page(overlay_page)
    #     except IndexError:
    #         pass  # No overlay for this page

    #     writer.add_page(target_page)

    # # Save the combined PDF with the employee's name
    # output_pdf_name = f"Workbook_{name}.pdf"
    # with open(output_pdf_name, "wb") as output_file:
    #     writer.write(output_file)

    # # Clean up
    # os.remove("data_overlay.pdf")
    # print(f"PDF generated: {output_pdf_name}")
