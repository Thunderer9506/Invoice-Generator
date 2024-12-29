import jinja2
import pdfkit
from datetime import datetime

# my_name = "Shaurya Srivastava"
# item1 = "TV"
# item2 = "Couch"
# item3 = "Washing machine"
# today_date = datetime.today().strftime("%d %b, %Y")
# context = {'my_name': my_name,'item1':item1,'item2':item2,'item3':item3,
#             'today_date':today_date,}

template_loader = jinja2.FileSystemLoader('./')
template_emv = jinja2.Environment(loader=template_loader)

template = template_emv.get_template('basic_template.html')
output_text = template.render()

config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
pdfkit.from_string(output_text, 'pdf_generated.pdf', configuration=config,css="style.css")