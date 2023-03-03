from PyPDF2 import PdfWriter, PdfReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# set font
pdfmetrics.registerFont(TTFont('F1', 'THSARABUNNEW.ttf'))



packet = io.BytesIO()
can = canvas.Canvas(packet, pagesize=A4)

# name
can.setFont('F1', 16)
can.drawString(150, 602, "เด็กหญิงสมหญิง จริงใจ")

# date
can.drawString(370, 602, "14")

# month
can.drawString(460, 602, "กันยายน")

# year budhist
can.drawString(93, 581, "2537")

# age
can.drawString(158, 581, "11")

# id_number
start_position = 314.5
for i in range(13):

    can.drawString(start_position, 579, '1')
    start_position += 18

# address
can.drawString(134, 561, "110")
can.drawString(193, 561, "110")
can.drawString(256, 561, "ทุ่งกุลา")
can.drawString(365, 561, "ท่าตูม")
can.drawString(468, 561, "สุรินทร์")



can.save()

#move to the beginning of the StringIO buffer
packet.seek(0)

# create a new PDF with Reportlab
new_pdf = PdfReader(packet)
# read your existing PDF
existing_pdf = PdfReader(open("subscription.pdf", "rb"))
output = PdfWriter()
# add the "watermark" (which is the new pdf) on the existing page
page = existing_pdf.pages[0]
page.merge_page(new_pdf.pages[0])
output.add_page(page)
# finally, write "output" to a real file
output_stream = open("destination.pdf", "wb")
output.write(output_stream)
output_stream.close()