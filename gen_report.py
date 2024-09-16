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
template_path = "MBA_template.pdf"
pdf_output_dir = "Generated PDFs"
term = 'Fall 2077'

# Ensure output directory exists
if not os.path.exists(pdf_output_dir):
    os.makedirs(pdf_output_dir)

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    name = row['Name']  # Using 'Name' column as the identifier
    print(index)
    save_chart(row, df)
    

    print(f"Generating PDF for {name}")

    # Create the overlay for each employee
    o = canvas.Canvas("data_overlay.pdf", pagesize=letter)
    o.setFont("Helvetica", 13)
    o.setFillColor("white")
    o.drawString(57, 534, name + ', ' + term)

    # Read the template PDF
    template = PdfReader(template_path)

    # Loop through pages of the template and overlay charts
    for pagenum in range(len(template.pages) - 1):
        o.showPage()

        if pagenum == 1:
            chart1 = ImageReader("CHART_IMAGES/E.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/A.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")

        if pagenum == 2:
            chart1 = ImageReader("CHART_IMAGES/C.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/ES.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")

        if pagenum == 3:
            chart1 = ImageReader("CHART_IMAGES/OTE.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/CSE.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")
        
        if pagenum == 4:
            chart1 = ImageReader("CHART_IMAGES/RTC.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
        
        if pagenum == 6:
            chart1 = ImageReader("CHART_IMAGES/Total_Size.png")
            o.drawImage(chart1, 160, 440, 300, 150, mask="auto")

        if pagenum == 7:
            chart1 = ImageReader("CHART_IMAGES/strength.png")
            o.drawImage(chart1, 95, 220, 400, 220, mask="auto")
        
        if pagenum == 8:
            chart1 = ImageReader("CHART_IMAGES/breadth.png")
            o.drawImage(chart1, 105, 270, 400, 220, mask="auto")

        if pagenum == 9:
            chart1 = ImageReader("CHART_IMAGES/PS.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/NA1.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")

        if pagenum == 10:
            chart1 = ImageReader("CHART_IMAGES/II.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/SA.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")

        if pagenum == 11:
            chart1 = ImageReader("CHART_IMAGES/AS.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")

        if pagenum == 13: #super
            chart1 = ImageReader("CHART_IMAGES/CR.png")
            o.drawImage(chart1, 160, 365, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/CR_super.png")
            o.drawImage(chart2, 160, 155, 300, 150, mask="auto")

        if pagenum == 14: #super
            chart1 = ImageReader("CHART_IMAGES/GTL.png")
            o.drawImage(chart1, 160, 365, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/GTL_super.png")
            o.drawImage(chart2, 160, 155, 300, 150, mask="auto")

        if pagenum == 15: #super
            chart1 = ImageReader("CHART_IMAGES/VBL.png")
            o.drawImage(chart1, 160, 369, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/VBL_super.png")
            o.drawImage(chart2, 160, 159, 300, 150, mask="auto")

        if pagenum == 16: #super
            chart1 = ImageReader("CHART_IMAGES/VL.png")
            o.drawImage(chart1, 160, 370, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/VL_super.png")
            o.drawImage(chart2, 160, 160, 300, 150, mask="auto")
        
        if pagenum == 17:
            chart1 = ImageReader("CHART_IMAGES/EmpL.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
        
        if pagenum == 18:
            chart1 = ImageReader("CHART_IMAGES/LBE.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/PDM.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")

        if pagenum == 19:
            chart1 = ImageReader("CHART_IMAGES/COACH.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/INF.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")

        if pagenum == 20:
            chart1 = ImageReader("CHART_IMAGES/ShowCon.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/PE.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")

        if pagenum == 21:
            chart1 = ImageReader("CHART_IMAGES/MEAN.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/COMP.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")
        
        if pagenum == 22:
            chart1 = ImageReader("CHART_IMAGES/SD.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/IMP.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")

        if pagenum == 24: #super
            chart1 = ImageReader("CHART_IMAGES/EL.png")
            o.drawImage(chart1, 160, 370, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/EL_super.png")
            o.drawImage(chart2, 160, 160, 300, 150, mask="auto")

        if pagenum == 25: #super
            chart1 = ImageReader("CHART_IMAGES/BLM.png")
            o.drawImage(chart1, 160, 370, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/BLM_super.png")
            o.drawImage(chart2, 160, 160, 300, 150, mask="auto")
        
        if pagenum == 26: #super
            chart1 = ImageReader("CHART_IMAGES/SE.png")
            o.drawImage(chart1, 160, 370, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/SE_super.png")
            o.drawImage(chart2, 160, 160, 300, 150, mask="auto")

        if pagenum == 27:
            chart1 = ImageReader("CHART_IMAGES/EM.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")

        if pagenum == 29:
            chart1 = ImageReader("CHART_IMAGES/PC.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/ORG.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")

        if pagenum == 30:
            chart1 = ImageReader("CHART_IMAGES/DOER.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/CHAL.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")

        if pagenum == 31:
            chart1 = ImageReader("CHART_IMAGES/INNOV.png")
            o.drawImage(chart1, 160, 460, 300, 150, mask="auto")
                
            chart2 = ImageReader("CHART_IMAGES/TB.png")
            o.drawImage(chart2, 160, 156, 300, 150, mask="auto")
        
        if pagenum == 32:
            chart1 = ImageReader("CHART_IMAGES/CONN.png")
            o.drawImage(chart1, 160, 455, 300, 150, mask="auto")

        o.setFont("Helvetica", 10)
        if pagenum not in (0, 5, 12, 23, 28, 33, 34, 35):
            o.drawString(560 - o.stringWidth(name.split(' ')[1]), 38, f"{name.split(' ')[1]} {pagenum}")

    o.save()

    # Read the overlay and merge with the template
    overlay = PdfReader("data_overlay.pdf")
    writer = PdfWriter()

    for pagenum in range(len(template.pages)):
        target_page = template.pages[pagenum]

        try:
            overlay_page = overlay.pages[pagenum]
            target_page.merge_page(overlay_page)
        except IndexError:
            pass  # No overlay for this page

        writer.add_page(target_page)

    # Save the combined PDF with the employee's name
    writer.write(f"Workbook_{name}.pdf")

    # Clean up
    os.remove("data_overlay.pdf")
    print(f"PDF generated: Workbook_{name}.pdf")
